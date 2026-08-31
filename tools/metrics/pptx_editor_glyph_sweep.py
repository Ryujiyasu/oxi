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

# One call per deck: lay out every slide and hand back the lines with the x of
# each character boundary, in points from the line's start.
JS = r"""
async (bytes) => {
  const m = await import('./face-metrics.js');
  const w = await import('./oxidocs_wasm.js');
  await w.default();
  const pres = w.parse_presentation(new Uint8Array(bytes));
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
  pres.slides.forEach((slide, si) => {
    for (const sh of slide.shapes) {
      const paras = sh.content?.TextBox?.paragraphs ?? sh.content?.AutoShape?.paragraphs;
      if (!paras || !paras.length) continue;
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
        // Per RUN, not per line: a paragraph that opens bold and continues
        // regular is two faces, and one of them is wrong for the other's half.
        const parts = (line.segments && line.segments.length) ? line.segments
          : [{ text: line.text, family: line.family, font_size: line.font_size,
               bold: line.bold, italic: line.italic }];
        const offs = [0];
        let x = 0, ok = true;
        for (const sg of parts) {
          const cps = [...sg.text];
          const em = m.measureFace(sg.family, sg.bold, sg.italic, sg.text)
            || cps.map(c => w.slide_face_advance(sg.family, sg.bold, sg.italic, c));
          if (!em || em.some(v => v === null || v === undefined)) { ok = false; break; }
          cps.forEach((c, i) => { x += em[i] * sg.font_size; offs.push(x); });
        }
        if (!ok) continue;
        out.push({ slide: si + 1, text: line.text, x: sh.x + line.x,
                   y: sh.y + line.baseline,
                   size: line.font_size, family: line.family,
                   bold: !!line.bold, offs });
      }
    }
  });
  return out;
}
"""


def pdf_pages(pdf: Path) -> dict[int, list]:
    """Every character of every page with its x and baseline, in points."""
    out: dict[int, list] = {}
    doc = pymupdf.open(pdf)
    for pno in range(len(doc)):
        chars = []
        for block in doc[pno].get_text("rawdict")["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    for ch in span.get("chars", []):
                        chars.append({"c": ch["c"], "x": ch["bbox"][0],
                                      "y": span["origin"][1]})
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
    on_baseline, anywhere = None, None
    want_y = line.get("y")
    for i, c in enumerate(chars):
        if c["c"] != line["text"][0]:
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
        if anywhere is None or abs(run[0]["x"] - line["x"]) < abs(anywhere[0]["x"] - line["x"]):
            anywhere = run
        if want_y is not None and abs(c["y"] - want_y) <= near:
            if on_baseline is None or (abs(run[0]["x"] - line["x"])
                                       < abs(on_baseline[0]["x"] - line["x"])):
                on_baseline = run
    if on_baseline is not None:
        return on_baseline, False
    return anywhere, True


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--decks", default="dev", choices=["dev", "all"])
    ap.add_argument("--limit", type=int, default=0, help="stop after N decks")
    ap.add_argument("--report", type=int, default=25, help="offending lines to list")
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

    os.chdir(REPO / "web")
    httpd = socketserver.TCPServer(("127.0.0.1", 0),
                                   http.server.SimpleHTTPRequestHandler)
    httpd.RequestHandlerClass.log_message = lambda *a, **k: None
    port = httpd.server_address[1]
    threading.Thread(target=httpd.serve_forever, daemon=True).start()

    offenders = []
    totals = {"lines": 0, "matched": 0, "unsure": 0}
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(f"http://127.0.0.1:{port}/pptx-editor.html")
        page.wait_for_function(
            "() => document.getElementById('status').textContent.includes('ready')",
            timeout=60000)
        for stem, pptx, pdf in pairs:
            try:
                lines = page.evaluate(JS, list(pptx.read_bytes()))
                pages = pdf_pages(pdf)
            except Exception as e:
                print(f"{stem:6} refused: {str(e)[:60]}", flush=True)
                continue
            worst, matched, unsure = 0.0, 0, 0
            for line in lines:
                chars = pages.get(line["slide"])
                if not chars:
                    continue
                run, loose = match_line(line, chars)
                if not run:
                    continue
                if loose:
                    unsure += 1
                    continue
                matched += 1
                d = max(abs(line["offs"][k] - (run[k]["x"] - run[0]["x"]))
                        for k in range(len(run)))
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
            print(f"{stem:6} {len(lines):5} lines  {matched:5} matched  "
                  f"{unsure:4} unsure  worst {worst:7.3f}pt", flush=True)
        browser.close()

    offenders.sort(key=lambda o: -o["worst"])
    print(f"\n{totals['matched']} of {totals['lines']} lines matched to a truth PDF; "
          f"{len(offenders)} disagree by more than 1pt")
    by_family: dict[str, int] = {}
    for o in offenders:
        by_family[o["family"]] = by_family.get(o["family"], 0) + 1
    if by_family:
        print("\nby face:")
        for fam, n in sorted(by_family.items(), key=lambda kv: -kv[1])[:12]:
            print(f"   {n:4}  {fam}")
    if offenders:
        print(f"\nworst {args.report}:")
        for o in offenders[: args.report]:
            print(f"   {o['deck']} s{o['slide']:<3} {o['worst']:8.3f}pt  "
                  f"{o['size']:6.2f}pt {o['family'][:22]:23}"
                  f"{'B' if o['bold'] else ' '} n={o['n']:3} {o['text']!r}")
    (REPO / "pipeline_data" / "pptx_editor_glyph_sweep.json").write_text(
        json.dumps({"totals": totals, "offenders": offenders[:400]}, indent=1),
        encoding="utf-8")


if __name__ == "__main__":
    main()
