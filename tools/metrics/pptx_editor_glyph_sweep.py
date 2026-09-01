"""Every line the browser engine lays out, against where PowerPoint drew it.

`pptx_editor_glyph_probe.py` asks this of one slide. This asks it of the whole
corpus, because the first time it was asked it found a 41pt defect immediately
(the layout took a face from the runs or the theme and never from the level).
A rule that is wrong is wrong on many lines, so the ones that disagree most are
the queue.

Output per deck: how many lines could be matched to the truth PDF, and the
worst per-character disagreement. The offenders are listed with their face and
size so a class of defect is visible as a class.

    python tools/metrics/pptx_editor_glyph_sweep.py [--decks dev|all] [--limit N]
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

# Parse once per deck, then ask for ONE SLIDE at a time.
#
# ★The whole deck in one message killed the page on d12 -- and a page that dies
# takes the driver connection with it, so every deck after it could not even
# launch a browser. One slide per call bounds what crosses the boundary, and a
# failure costs one slide.
PARSE_JS = r"""
async (url) => {
  const m = await import('./face-metrics.js');
  const w = await import('./oxidocs_wasm.js');
  await w.default();
  window.__m = m;
  window.__w = w;
  // ★Fetched, not passed in. Handing the file to `page.evaluate` marshals it
  // as a JSON array of numbers, and d12 is 40MB -- forty million of them,
  // which closed the driver connection and looked exactly like a parser crash.
  const res = await fetch(url);
  const buf = new Uint8Array(await res.arrayBuffer());
  window.__pres = w.parse_presentation(buf);
  return window.__pres.slides.length;
}
"""

SLIDE_JS = r"""
(si) => {
  const m = window.__m, w = window.__w, pres = window.__pres;
  const st = pres.master_styles || {};
  const fallback = pres.minor_font || 'Calibri';
  const levelsOf = (sh) => {
    if (sh.ph_levels && sh.ph_levels.length) return sh.ph_levels;
    const t = (sh.ph_type || '').toLowerCase();
    if (t === 'title' || t === 'ctrtitle') return st.title || [];
    if (t) return st.body || [];
    return st.other || [];
  };
  const out = [];
  let skipped = 0;
  for (const sh of pres.slides[si].shapes) {
    const paras = sh.content?.TextBox?.paragraphs ?? sh.content?.AutoShape?.paragraphs;
    if (!paras || !paras.length) continue;
    // ★A rotated shape's lines run along an axis this layout does not turn:
    // the engine answers in the shape's own frame, so comparing its offsets
    // against the PDF's x measures the rotation, not the advances. d06 slide
    // 34's 'HIGH VALUE 1' sits in a shape at +90 and read as 48pt of
    // horizontal "defect" and 29pt of vertical.
    if (sh.rotation) { skipped++; continue; }
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
      // ★The ENGINE's placement, which is what the editor draws: each advance
      // on the master unit. Accumulating the exact design advance here instead
      // measured a model the editor no longer uses -- and one PowerPoint never
      // used, per `read_pptx_drawgrid_com.py`.
      //
      // Per RUN, not per line: a paragraph that opens bold and continues
      // regular is two faces, and one of them is wrong for the other's half.
      const parts = (line.segments && line.segments.length) ? line.segments
        : [{ text: line.text, family: line.family, font_size: line.font_size,
             bold: line.bold, italic: line.italic }];
      let offs = null, ok = true;
      try { offs = w.slide_glyph_offsets(line, adv); } catch (e) { offs = null; }
      if (!offs) {
        offs = [0];
        let x = 0;
        for (const sg of parts) {
          const cps = [...sg.text];
          const em = m.measureFace(sg.family, sg.bold, sg.italic, sg.text)
            || cps.map(c => w.slide_face_advance(sg.family, sg.bold, sg.italic, c));
          if (!em || em.some(v => v === null || v === undefined)) { ok = false; break; }
          cps.forEach((c, i) => { x += em[i] * sg.font_size; offs.push(x); });
        }
      }
      if (!ok) continue;
      out.push({ slide: si + 1, text: line.text, x: sh.x + line.x,
                 y: sh.y + line.baseline,
                 size: line.font_size, family: line.family,
                 bold: !!line.bold, offs });
    }
  }
  return { lines: out, skipped };
}
"""


def deck_url(port: int, pptx: Path) -> str:
    """The deck's URL under the repo-rooted server."""
    rel = pptx.resolve().relative_to(REPO).as_posix()
    return f"http://127.0.0.1:{port}/{rel}"


def pdf_pages(pdf: Path) -> dict[int, list]:
    """Every character of every page with its x, baseline and own width.

    The width is carried because a LIGATURE is one glyph that the text layer
    reports as two characters: PowerPoint drew `fl` as a single glyph, and the
    extraction gives 'f' the pair's whole advance and 'l' a width of ZERO, at
    the position the following letter starts. Reading that as a character
    position says the engine is 2.686pt out when it is not -- d05 slide 4 draws
    the same sentence three times, two of them without the ligature and
    matching to 0.224pt.
    """
    out: dict[int, list] = {}
    doc = pymupdf.open(pdf)
    for pno in range(len(doc)):
        chars = []
        for block in doc[pno].get_text("rawdict")["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    for ch in span.get("chars", []):
                        chars.append({"c": ch["c"], "x": ch["bbox"][0],
                                      "w": ch["bbox"][2] - ch["bbox"][0],
                                      "y": span["origin"][1],
                                      "size": span["size"]})
        out[pno + 1] = chars
    doc.close()
    return out


def match_line(line, chars, near=6.0):
    """The PDF characters that make up `line`, or None.

    ★Identity before measurement. A slide can carry the same words several
    times -- d06 slide 34 has seven lines starting with 'C' -- and picking the
    candidate that merely starts nearest in x paired one 'Competitor' with
    another's characters and reported 103pt of "disagreement" that was entirely
    the instrument's. So a candidate must sit on the baseline the engine placed
    the line on, and only then is the nearest x taken.

    Returns (run, loose) where `loose` says no candidate was on the baseline
    and the answer is a guess -- those are counted apart from real defects.
    """
    on_baseline, anywhere, seen = None, None, 0
    want_y = line.get("y")
    want_size = line.get("size")
    for i, c in enumerate(chars):
        if c["c"] != line["text"][0]:
            continue
        # ★Size is part of the identity. d32 slide 21 has a 284pt '!' in a
        # decorative shape and a 20pt '!' in body text; the deck draws only the
        # small one, the text was "unique" on the page, and pairing them
        # reported 123pt of vertical "defect" that was the instrument's.
        if want_size and abs(c.get("size", want_size) - want_size) > 0.1 * want_size:
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
        if not run:
            continue
        seen += 1
        if anywhere is None or abs(run[0]["x"] - line["x"]) < abs(anywhere[0]["x"] - line["x"]):
            anywhere = run
        if want_y is not None and abs(c["y"] - want_y) <= near:
            if on_baseline is None or (abs(run[0]["x"] - line["x"])
                                       < abs(on_baseline[0]["x"] - line["x"])):
                on_baseline = run
    # ★One candidate means the identity is certain, wherever it sits -- which
    # is the only case where a BASELINE difference can be read as a defect
    # rather than as the instrument having paired the wrong two lines. That is
    # what d06's "194pt vertical error" was: 34 lines whose text repeats across
    # the deck, paired with each other's characters.
    if seen == 1:
        return anywhere, False, True
    if on_baseline is not None:
        return on_baseline, False, False
    return anywhere, True, False


# How far a pairing's first character may sit from where the engine put the
# line before the pairing stops being evidence about ADVANCES.
#
# ★d30 slide 24 reported 10.698pt -- the worst line in the whole corpus -- and
# none of it was the engine's. Its paragraphs begin with a literal space, and
# PowerPoint draws the autonum marker's trailing tab as a space glyph right
# after "2."; the two spaces are the same character, both followed by "Lorem",
# so the sequence matched one character early and every advance after it
# carried the tab's width. The engine had the text at marL exactly, which is
# where PowerPoint put it. A run that starts somewhere else is a different run
# -- or a position defect, which is worth reporting AS one rather than folding
# into a number about advances.
POSITION_NEAR = 2.0


def horizontal_offset(line, run):
    """How far the pairing's first character is from where the engine put it.

    The PDF's x is an ink bbox and the engine's is a pen position, so a few
    tenths of a point of left side bearing are expected and mean nothing.
    """
    if not run or line.get("x") is None:
        return None
    return run[0]["x"] - line["x"]


def vertical_error(line, run):
    """How far the engine's baseline is from the one PowerPoint drew on.

    ★Only asked of a line whose text appears ONCE on the page. A line that
    misses the baseline window drops out of the horizontal measurement as
    "unsure", which looks like instrument noise and is not -- but reading a
    baseline difference off an ambiguous match measures the pairing, not the
    layout. Uniqueness is the identity that makes the number mean something.
    """
    if not run or line.get("y") is None:
        return None
    return run[0]["y"] - line["y"]


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--decks", default="dev", choices=["dev", "all"])
    ap.add_argument("--limit", type=int, default=0, help="stop after N decks")
    ap.add_argument("--report", type=int, default=25, help="offending lines to list")
    ap.add_argument("--fresh", action="store_true", help="ignore what was already measured")
    args = ap.parse_args()

    bench = REPO / "pipeline_data" / "pptx_benchmark"
    pairs = []
    roots = [(bench / "dev" / "pptx", bench / "dev" / "pdf")]
    if args.decks == "all":
        roots.append((bench / "pptx", bench / "ssim_pptx" / "ppt_pdf"))
    for pptx_dir, pdf_dir in roots:
        for pptx in sorted(pptx_dir.glob("*.pptx")):
            stem = pptx.name.split("__")[0]
            pdf = next(iter(sorted(pdf_dir.glob(stem + "*.pdf"))), None)
            if pdf:
                pairs.append((stem, pptx, pdf))
    if args.limit:
        pairs = pairs[: args.limit]

    # Resumable: one JSON line per deck, so a run that is interrupted (or a
    # machine that is busy) picks up where it stopped instead of starting over.
    store = REPO / "pipeline_data" / "pptx_editor_glyph_sweep.jsonl"
    done: dict[str, dict] = {}
    if store.exists() and not args.fresh:
        for row in store.read_text(encoding="utf-8").splitlines():
            if row.strip():
                r = json.loads(row)
                done[r["deck"]] = r

    # Served from the repo root so the page can fetch a deck by URL.
    os.chdir(REPO)
    httpd = socketserver.TCPServer(("127.0.0.1", 0),
                                   http.server.SimpleHTTPRequestHandler)
    httpd.RequestHandlerClass.log_message = lambda *a, **k: None
    port = httpd.server_address[1]
    threading.Thread(target=httpd.serve_forever, daemon=True).start()

    offenders = []
    vertical = []
    misplaced = []
    totals = {"lines": 0, "matched": 0, "unsure": 0, "misplaced": 0}
    # ★A deck can take the page down with it -- d12 closed the driver
    # connection mid-evaluate, and every deck after it then "refused" against a
    # browser that was no longer there. A failure has to cost one deck, not the
    # rest of the run, so the page is rebuilt and the deck retried once.

    for r in done.values():
        for k in totals:
            totals[k] += r["totals"].get(k, 0)
        offenders += r["offenders"]
        vertical += r.get("vertical", [])
        misplaced += r.get("misplaced", [])
    with sync_playwright() as p:
        state = {"browser": None, "page": None}

        def fresh_page():
            if state["browser"] is not None:
                try:
                    state["browser"].close()
                except Exception:
                    pass
            state["browser"] = p.chromium.launch()
            state["page"] = state["browser"].new_page()
            state["page"].goto(f"http://127.0.0.1:{port}/web/pptx-editor.html")
            state["page"].wait_for_function(
                "() => document.getElementById('status').textContent.includes('ready')",
                timeout=60000)

        fresh_page()
        for stem, pptx, pdf in pairs:
            if stem in done:
                d = done[stem]
                print(f"{stem:6} {d['totals']['lines']:5} lines  "
                      f"{d['totals']['matched']:5} matched  "
                      f"{d['totals']['unsure']:4} unsure  "
                      f"worst dx {d['worst']:7.3f}pt  "
                      f"dy {d.get('worst_dy', 0.0):7.3f}pt   (stored)", flush=True)
                continue
            lines, turned, n_slides = [], 0, None
            for attempt in (1, 2):
                try:
                    n_slides = state["page"].evaluate(PARSE_JS, deck_url(port, pptx))
                    break
                except Exception as e:
                    print(f"{stem:6} attempt {attempt} refused: {str(e)[:52]}",
                          flush=True)
                    try:
                        fresh_page()
                    except Exception as e2:
                        print(f"       could not restart the browser: {str(e2)[:50]}",
                              flush=True)
                        break
            if n_slides is None:
                continue
            lost = 0
            for si in range(n_slides):
                try:
                    got = state["page"].evaluate(SLIDE_JS, si)
                except Exception:
                    lost += 1
                    try:
                        fresh_page()
                        state["page"].evaluate(PARSE_JS, deck_url(port, pptx))
                    except Exception:
                        break
                    continue
                lines += got["lines"]
                turned += got["skipped"]
            if lost:
                print(f"{stem:6} {lost} slides could not be laid out", flush=True)
            try:
                pages = pdf_pages(pdf)
            except Exception as e:
                print(f"{stem:6} no truth PDF: {str(e)[:50]}", flush=True)
                continue
            worst, matched, unsure = 0.0, 0, 0
            dys = []
            for line in lines:
                chars = pages.get(line["slide"])
                if not chars:
                    continue
                run, loose, certain = match_line(line, chars)
                if not run:
                    continue
                dy = vertical_error(line, run) if certain else None
                if dy is not None:
                    dys.append(abs(dy))
                    if abs(dy) > 1.0:
                        vertical.append({
                            "deck": stem, "slide": line["slide"],
                            "dy": round(dy, 3), "size": line["size"],
                            "family": line["family"], "text": line["text"][:38]})
                if loose:
                    unsure += 1
                    continue
                dx = horizontal_offset(line, run)
                if dx is not None and abs(dx) > POSITION_NEAR:
                    misplaced.append({
                        "deck": stem, "slide": line["slide"],
                        "dx": round(dx, 3), "size": line["size"],
                        "family": line["family"], "text": line["text"][:38]})
                    continue
                matched += 1
                # A zero-width character is the tail of a ligature: one glyph
                # was drawn for two characters, and the second has no position
                # of its own to be right or wrong about. The characters either
                # side still answer the question.
                real = [k for k in range(len(run)) if run[k].get("w", 1.0) > 0.01]
                if not real:
                    continue
                d = max(abs(line["offs"][k] - (run[k]["x"] - run[0]["x"]))
                        for k in real)
                worst = max(worst, d)
                if d > 1.0:
                    offenders.append({
                        "deck": stem, "slide": line["slide"], "worst": round(d, 3),
                        "size": line["size"], "family": line["family"],
                        "bold": line["bold"], "n": len(run),
                        "text": line["text"][:38]})
            totals["lines"] += len(lines)
            totals["matched"] += matched
            totals["unsure"] += unsure
            off_here = [m for m in misplaced if m["deck"] == stem]
            totals["misplaced"] += len(off_here)
            worst_dy = max(dys) if dys else 0.0
            mine = [o for o in offenders if o["deck"] == stem]
            with store.open("a", encoding="utf-8") as fh:
                fh.write(json.dumps({
                    "deck": stem, "worst": round(worst, 3),
                    "worst_dy": round(worst_dy, 3),
                    "totals": {"lines": len(lines), "matched": matched,
                               "unsure": unsure, "misplaced": len(off_here)},
                    "offenders": mine,
                    "misplaced": off_here,
                    "vertical": [v for v in vertical if v["deck"] == stem]})
                    + chr(10))
            print(f"{stem:6} {len(lines):5} lines  {matched:5} matched  "
                  f"{unsure:4} unsure  {len(off_here):3} misplaced  "
                  f"{turned:3} turned  "
                  f"worst dx {worst:7.3f}pt  dy {worst_dy:7.3f}pt", flush=True)
        if state["browser"] is not None:
            # A browser that already died takes the close with it; the results
            # are in hand by now, so a shutdown failure must not lose them.
            try:
                state["browser"].close()
            except Exception:
                pass

    offenders.sort(key=lambda o: -o["worst"])
    misplaced.sort(key=lambda m: -abs(m["dx"]))
    print(f"\n{totals['matched']} of {totals['lines']} lines matched to a truth PDF; "
          f"{len(offenders)} disagree by more than 1pt")
    if misplaced:
        print(f"\n{len(misplaced)} pairings begin more than {POSITION_NEAR}pt from "
              f"where the engine put the line, so they measure POSITION and not "
              f"advances; worst {args.report // 2}:")
        for m in misplaced[: args.report // 2]:
            print(f"   {m['deck']} s{m['slide']:<3} {m['dx']:+8.3f}pt  "
                  f"{m['size']:6.2f}pt {m['family'][:22]:23} {m['text']!r}")
    by_family: dict[str, int] = {}
    for o in offenders:
        by_family[o["family"]] = by_family.get(o["family"], 0) + 1
    if by_family:
        print("\nby face:")
        for fam, n in sorted(by_family.items(), key=lambda kv: -kv[1])[:12]:
            print(f"   {n:4}  {fam}")
    vertical.sort(key=lambda v: -abs(v["dy"]))
    if vertical:
        print(f"\n{len(vertical)} lines sit more than 1pt off the baseline "
              f"PowerPoint drew on; worst {args.report // 2}:")
        for v in vertical[: args.report // 2]:
            print(f"   {v['deck']} s{v['slide']:<3} {v['dy']:+8.3f}pt  "
                  f"{v['size']:6.2f}pt {v['family'][:22]:23} {v['text']!r}")
    if offenders:
        print(f"\nworst {args.report} by advance:")
        for o in offenders[: args.report]:
            print(f"   {o['deck']} s{o['slide']:<3} {o['worst']:8.3f}pt  "
                  f"{o['size']:6.2f}pt {o['family'][:22]:23}"
                  f"{'B' if o['bold'] else ' '} n={o['n']:3} {o['text']!r}")
    (REPO / "pipeline_data" / "pptx_editor_glyph_sweep.json").write_text(
        json.dumps({"totals": totals, "offenders": offenders[:400],
                    "misplaced": misplaced[:400]}, indent=1),
        encoding="utf-8")


if __name__ == "__main__":
    main()
