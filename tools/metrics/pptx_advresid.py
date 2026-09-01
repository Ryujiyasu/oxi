"""The per-character residual of one line, with its shape, not just its worst.

`pptx_editor_glyph_sweep.py` says a line's worst character is 2.461pt out. That
number cannot tell a SCALE error (every advance a fixed fraction too wide) from
a ROUNDING walk (each glyph placed on a grid) from a SQUEEZE (a few pieces of
the line pulled in). Those want different fixes, so the residual has to be read
along the line rather than maximised over it.

For each character the probe prints where the engine puts it, where the truth
PDF puts it, and the difference -- and then scores these models against the
same line:

    exact     cumulative sum of em * size, which is what the engine does
    grid      cumulative sum of each advance rounded to 1/Npt, swept over N so
              the unit is measured rather than assumed. 1/8 is the break
              model's own unit (`pptx-master-unit-break-law`)
    scaled    exact, times the single factor that best fits this line

★What this found (2026-09-01), and why it did NOT become an implementation:
1/8pt is uniquely the best grid over the dev corpus -- 2796 lines, mean worst
0.276 -> 0.240pt, lines over 1pt 47 -> 20, and every OTHER grid including finer
ones is worse than exact, which is the shape a real unit makes. But it holds
only for INTEGER point sizes (0.277 -> 0.237) and reverses at fractional ones
(0.264 -> 0.284), and `gen_pptx_drawgrid.py`'s minimal repro shows PowerPoint's
own per-glyph steps alternating between two values that are neither model's.
A partial law is a wrong law (`no EXCEPTION stacking`), so this stays an
instrument until the alternation is explained.

    python tools/metrics/pptx_advresid.py --deck d35 --slide 25 [--min 1.0] [--chars]
    python tools/metrics/pptx_advresid.py --corpus
"""
from __future__ import annotations

import argparse
import http.server
import json
import os
import socketserver
import sys
import threading
from pathlib import Path

import pymupdf
from playwright.sync_api import sync_playwright

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]

# Parse once per deck (see `pptx_editor_glyph_sweep.py`: the whole deck in one
# message kills the page, and re-parsing per slide costs the parse 40 times).
PARSE_JS = r"""
async (url) => {
  const m = await import('./face-metrics.js');
  const w = await import('./oxidocs_wasm.js');
  await w.default();
  window.__m = m;
  window.__w = w;
  const res = await fetch(url);
  window.__pres = w.parse_presentation(new Uint8Array(await res.arrayBuffer()));
  return window.__pres.slides.length;
}
"""

# The sweep's own layout call, returning the per-character EM advances as well
# as the accumulated offsets so the models below can be built from the same
# numbers the editor drew with.
JS = r"""
(slideNo) => {
  const m = window.__m, w = window.__w, pres = window.__pres;
  const st = pres.master_styles || {};
  const fallback = pres.minor_font || 'Calibri';
  const slide = pres.slides[slideNo - 1];
  if (!slide) return [];
  const levelsOf = (sh) => {
    if (sh.ph_levels && sh.ph_levels.length) return sh.ph_levels;
    const t = (sh.ph_type || '').toLowerCase();
    if (t === 'title' || t === 'ctrtitle') return st.title || [];
    if (t) return st.body || [];
    return st.other || [];
  };
  const out = [];
  for (const sh of slide.shapes) {
    const paras = sh.content?.TextBox?.paragraphs ?? sh.content?.AutoShape?.paragraphs;
    if (!paras || !paras.length) continue;
    if (sh.rotation) continue;
    const lv = levelsOf(sh);
    const runs = [];
    paras.forEach(p => {
      const inherited = lv.length
        ? (lv[Math.min(p.lvl || 0, lv.length - 1)] || {}).font_family : null;
      p.runs.forEach(r => runs.push({
        text: r.text, font_family: r.font_family || inherited || fallback,
        bold: r.bold, italic: r.italic }));
    });
    const adv = m.collectAdvances(runs, fallback);
    let lay = null;
    try {
      lay = w.layout_slide_shape(sh, paras, lv, sh.ph_levels || [], fallback, adv);
    } catch (e) { continue; }
    if (!lay || !lay.complete) continue;
    for (const line of lay.lines) {
      if (!line.text.trim()) continue;
      const parts = (line.segments && line.segments.length) ? line.segments
        : [{ text: line.text, family: line.family, font_size: line.font_size,
             bold: line.bold, italic: line.italic }];
      const ems = [], sizes = [];
      let ok = true;
      for (const sg of parts) {
        const cps = [...sg.text];
        const em = m.measureFace(sg.family, sg.bold, sg.italic, sg.text)
          || cps.map(ch => w.slide_face_advance(sg.family, sg.bold, sg.italic, ch));
        if (!em || em.some(v => v === null || v === undefined)) { ok = false; break; }
        cps.forEach((ch, i) => { ems.push(em[i]); sizes.push(sg.font_size); });
      }
      if (!ok) continue;
      out.push({ text: line.text, x: sh.x + line.x, baseline: sh.y + line.baseline,
                 size: line.font_size, family: line.family, bold: !!line.bold,
                 ems, sizes });
    }
  }
  return out;
}
"""


def pdf_pages(pdf: Path) -> dict[int, list]:
    """Every character of every page, in points. Opened once per deck."""
    doc = pymupdf.open(pdf)
    out: dict[int, list] = {}
    for pno in range(len(doc)):
        out[pno + 1] = _page_chars(doc[pno])
    doc.close()
    return out


def _page_chars(page):
    out = []
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                for ch in span.get("chars", []):
                    # ★The PEN position, not the ink box. They agree for
                    # 7428 of d35's 7647 characters and differ by a constant
                    # 8.59pt for the 219 in one transformed span, so reading
                    # the box measures that span's transform.
                    out.append({"c": ch["c"], "x": ch["origin"][0],
                                "w": ch["bbox"][2] - ch["bbox"][0],
                                "y": span["origin"][1], "size": span["size"]})
    return out


def match_line(line, chars):
    """The PDF run for `line`, nearest in x among the candidates on a baseline.

    The identity rules of the sweep are not repeated here: this probe is
    pointed at ONE line already known to disagree, and prints enough for the
    pairing to be checked by eye.
    """
    best = None
    for i, c in enumerate(chars):
        if c["c"] != line["text"][0]:
            continue
        if abs(c["size"] - line["size"]) > 0.1 * line["size"]:
            continue
        run = [c]
        j = i + 1
        for want in line["text"][1:]:
            while j < len(chars) and abs(chars[j]["y"] - c["y"]) > 1.5:
                j += 1
            if j >= len(chars) or chars[j]["c"] != want:
                run = None
                break
            run.append(chars[j])
            j += 1
        if run and (best is None
                    or abs(run[0]["x"] - line["x"]) < abs(best[0]["x"] - line["x"])):
            best = run
    return best


# Candidate placement grids, as divisions of a point. 8 is the break model's
# own unit (`pptx-master-unit-break-law`); the others are here so the fit is
# not credited to a unit that was assumed rather than measured.
GRIDS = (4, 8, 12, 16, 20, 32, 64)


def models(line, per_pt: float = 8.0):
    """Cumulative offsets under each draw model, in points."""
    exact, master = [0.0], [0.0]
    x = mu = 0.0
    for em, size in zip(line["ems"], line["sizes"]):
        x += em * size
        exact.append(x)
        # The break model's unit, asked as a placement unit.
        mu += round(em * size * per_pt) / per_pt
        master.append(mu)
    return exact, master


def corpus(args) -> None:
    """Both models against every line of every dev deck.

    One line agreeing is a coincidence; the question is whether the grid is
    better EVERYWHERE, and in particular whether it is ever worse -- a draw
    model that helps the long lines and hurts the short ones is not a law.
    """
    bench = REPO / "pipeline_data" / "pptx_benchmark" / "dev"
    pairs = []
    for pptx in sorted((bench / "pptx").glob("*.pptx")):
        stem = pptx.name.split("__")[0]
        pdf = next(iter(sorted((bench / "pdf").glob(stem + "*.pdf"))), None)
        if pdf:
            pairs.append((stem, pptx, pdf))
    if args.limit:
        pairs = pairs[: args.limit]

    os.chdir(REPO)
    httpd = socketserver.TCPServer(("127.0.0.1", 0),
                                   http.server.SimpleHTTPRequestHandler)
    httpd.RequestHandlerClass.log_message = lambda *a, **k: None
    port = httpd.server_address[1]
    threading.Thread(target=httpd.serve_forever, daemon=True).start()

    rows = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(f"http://127.0.0.1:{port}/web/pptx-editor.html")
        page.wait_for_function(
            "() => document.getElementById('status').textContent.includes('ready')",
            timeout=60000)
        for stem, pptx, pdf in pairs:
            url = f"http://127.0.0.1:{port}/{pptx.resolve().relative_to(REPO).as_posix()}"
            try:
                pages = pdf_pages(pdf)
                n_slides = page.evaluate(PARSE_JS, url)
            except Exception as e:
                print(f"{stem:6} refused: {str(e)[:50]}", flush=True)
                continue
            n_here = 0
            for sno in range(1, n_slides + 1):
                try:
                    lines = page.evaluate(JS, sno)
                except Exception:
                    continue
                chars = pages.get(sno)
                if not lines or not chars:
                    continue
                for line in lines:
                    run = match_line(line, chars)
                    if not run or len(run) < 4:
                        continue
                    truth = [c["x"] - run[0]["x"] for c in run]
                    exact, master = models(line)
                    n = min(len(truth), len(exact))
                    real = [k for k in range(n) if k == 0 or run[k]["w"] > 0.01]
                    if len(real) < 4:
                        continue
                    row = {
                        "deck": stem, "slide": sno, "n": len(run),
                        "family": line["family"], "size": line["size"],
                        "exact": max(abs(exact[k] - truth[k]) for k in real),
                        "master": max(abs(master[k] - truth[k]) for k in real),
                        "text": line["text"][:38],
                    }
                    for g in GRIDS:
                        _, m2 = models(line, float(g))
                        row[f"g{g}"] = max(abs(m2[k] - truth[k]) for k in real)
                    rows.append(row)
                    n_here += 1
            print(f"{stem:6} {n_here:5} lines", flush=True)
        browser.close()

    store = REPO / "pipeline_data" / "pptx_advresid.jsonl"
    with store.open("w", encoding="utf-8") as fh:
        for r in rows:
            print(json.dumps(r, ensure_ascii=False), file=fh)
    print()
    print(f"rows written to {store}")
    if not rows:
        print("no lines matched")
        return
    better = [r for r in rows if r["master"] < r["exact"] - 0.01]
    worse = [r for r in rows if r["master"] > r["exact"] + 0.01]
    print(f"\n{len(rows)} lines matched over {len({r['deck'] for r in rows})} decks")
    print(f"  worst character, exact model : {max(r['exact'] for r in rows):7.3f}pt"
          f"   mean {sum(r['exact'] for r in rows) / len(rows):.3f}pt")
    print(f"  worst character, master model: {max(r['master'] for r in rows):7.3f}pt"
          f"   mean {sum(r['master'] for r in rows) / len(rows):.3f}pt")
    print(f"  better on the grid: {len(better)}   worse: {len(worse)}   "
          f"same: {len(rows) - len(better) - len(worse)}")
    over = lambda rs, t: sum(1 for r in rs if r[t] > 1.0)  # noqa: E731
    print(f"  lines over 1.0pt: exact {over(rows, 'exact')}  "
          f"master {over(rows, 'master')}")
    print()
    print("  the grid itself, as divisions of a point:")
    print(f"   {'grid':>6}  {'mean worst':>10}  {'max worst':>10}  {'lines >1pt':>10}")
    print(f"   {'exact':>6}  {sum(r['exact'] for r in rows) / len(rows):10.3f}  "
          f"{max(r['exact'] for r in rows):10.3f}  {over(rows, 'exact'):10}")
    for g in GRIDS:
        k = f"g{g}"
        print(f"   {'1/' + str(g):>6}  {sum(r[k] for r in rows) / len(rows):10.3f}  "
              f"{max(r[k] for r in rows):10.3f}  {over(rows, k):10}")
    if worse:
        print("\n  worst regressions on the grid:")
        for r in sorted(worse, key=lambda r: r["exact"] - r["master"])[:12]:
            print(f"   {r['deck']} s{r['slide']:<3} {r['family'][:18]:19} "
                  f"{r['size']:5.1f}pt n={r['n']:3}  exact {r['exact']:6.3f} -> "
                  f"master {r['master']:6.3f}  {r['text']!r}")
    print("\n  worst remaining on the grid:")
    for r in sorted(rows, key=lambda r: -r["master"])[:12]:
        print(f"   {r['deck']} s{r['slide']:<3} {r['family'][:18]:19} "
              f"{r['size']:5.1f}pt n={r['n']:3}  exact {r['exact']:6.3f} -> "
              f"master {r['master']:6.3f}  {r['text']!r}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", default="d35")
    ap.add_argument("--slide", type=int, default=25)
    ap.add_argument("--min", type=float, default=1.0,
                    help="only report lines whose worst character exceeds this")
    ap.add_argument("--chars", action="store_true", help="print every character")
    ap.add_argument("--corpus", action="store_true",
                    help="score both models on every line of every dev deck")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()
    if args.corpus:
        corpus(args)
        return

    pptx = next(iter(sorted(
        (REPO / "pipeline_data/pptx_benchmark/dev/pptx").glob(args.deck + "*.pptx"))), None)
    pdf = next(iter(sorted(
        (REPO / "pipeline_data/pptx_benchmark/dev/pdf").glob(args.deck + "*.pdf"))), None)
    if not pptx or not pdf:
        sys.exit(f"no deck/pdf for {args.deck}")

    os.chdir(REPO)
    httpd = socketserver.TCPServer(("127.0.0.1", 0),
                                   http.server.SimpleHTTPRequestHandler)
    httpd.RequestHandlerClass.log_message = lambda *a, **k: None
    port = httpd.server_address[1]
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    url = f"http://127.0.0.1:{port}/{pptx.resolve().relative_to(REPO).as_posix()}"
    with sync_playwright() as p:
        b = p.chromium.launch()
        page = b.new_page()
        page.goto(f"http://127.0.0.1:{port}/web/pptx-editor.html")
        page.wait_for_function(
            "() => document.getElementById('status').textContent.includes('ready')",
            timeout=60000)
        page.evaluate(PARSE_JS, url)
        lines = page.evaluate(JS, args.slide)
        b.close()

    chars = pdf_pages(pdf).get(args.slide, [])
    print(f"{args.deck} slide {args.slide}: {len(lines)} engine lines, "
          f"{len(chars)} PDF characters\n")
    for line in lines:
        run = match_line(line, chars)
        if not run:
            continue
        truth = [c["x"] - run[0]["x"] for c in run]
        exact, master = models(line)
        n = min(len(truth), len(exact))
        # A zero-width character is a ligature's tail and has no position of
        # its own to be right or wrong about.
        real = [k for k in range(n) if k == 0 or run[k]["w"] > 0.01]
        worst = max(abs(exact[k] - truth[k]) for k in real)
        if worst < args.min:
            continue

        # The single factor that best fits this line, and what is left after it.
        num = sum(exact[k] * truth[k] for k in real)
        den = sum(exact[k] * exact[k] for k in real) or 1.0
        f = num / den
        scaled_worst = max(abs(exact[k] * f - truth[k]) for k in real)
        master_worst = max(abs(master[k] - truth[k]) for k in real)

        print(f"{line['size']:6.2f}pt {line['family'][:20]:21}"
              f"{'B' if line['bold'] else ' '} n={len(run):3}  {line['text'][:44]!r}")
        print(f"    exact  worst {worst:7.3f}pt")
        print(f"    master worst {master_worst:7.3f}pt   (each advance on a 1/8pt grid)")
        print(f"    scaled worst {scaled_worst:7.3f}pt   (factor {f:.6f}, "
              f"{(f - 1.0) * 1e4:+.1f} per 10k)")
        # Where the line loses it: the biggest per-character steps of the
        # residual, which is what tells a walk from a squeeze.
        steps = sorted(
            ((abs(truth[k] - exact[k] - (truth[k - 1] - exact[k - 1])), k)
             for k in real if k > 0), reverse=True)[:6]
        print("    largest steps: " + "  ".join(
            f"{line['text'][k-1]!r}{'->' if False else ''}{d:+.3f}" for d, k in steps))
        if args.chars:
            print("        k  ch   our adv  truth adv   on 1/8pt   truth/size (em)")
            for k in range(1, min(len(truth), len(exact))):
                ours = exact[k] - exact[k - 1]
                got = truth[k] - truth[k - 1]
                grid = round(ours * 8.0) / 8.0
                print(f"      {k:3} {line['text'][k-1]!r:4} {ours:9.4f} {got:10.4f} "
                      f"{grid:10.4f}   {got / line['size']:8.5f}"
                      f"{'   <- off grid' if abs(got - grid) > 0.02 else ''}")
        print()


if __name__ == "__main__":
    main()
