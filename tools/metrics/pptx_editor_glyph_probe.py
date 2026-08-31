"""Where the editor puts each glyph, against where PowerPoint put it.

The canvas editor now places characters the way the renderer does -- the exact
design advance accumulated, no kerning -- instead of letting `fillText` shape
the whole string. This asks the only question that settles whether that is
right: for the same line, does the editor's Nth character start where
PowerPoint's PDF says it starts?

Slide coordinates and PDF coordinates are both in points, so the comparison is
direct: `shape.x + line.x + offset[n]` against the PDF character's x.
"""
import argparse
import http.server
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

JS = r"""
async ([bytes, slideNo]) => {
  const m = await import('./face-metrics.js');
  const w = await import('./oxidocs_wasm.js');
  await w.default();
  const pres = w.parse_presentation(new Uint8Array(bytes));
  const st = pres.master_styles || {};
  const fallback = pres.minor_font || 'Calibri';
  const slide = pres.slides[slideNo - 1];
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
      // The editor's own chain, per RUN: what the page measured, else the
      // tables. A paragraph that opens bold and continues regular is two
      // faces, and either one is wrong for the other's half.
      const parts = (line.segments && line.segments.length) ? line.segments
        : [{ text: line.text, family: line.family, font_size: line.font_size,
             bold: line.bold, italic: line.italic }];
      const offs = [0];
      let x = 0, ok = true;
      for (const sg of parts) {
        const cps = [...sg.text];
        const em = m.measureFace(sg.family, sg.bold, sg.italic, sg.text)
          || cps.map(ch => w.slide_face_advance(sg.family, sg.bold, sg.italic, ch));
        if (!em || em.some(v => v === null || v === undefined)) { ok = false; break; }
        cps.forEach((ch, i) => { x += em[i] * sg.font_size; offs.push(x); });
      }
      if (!ok) continue;
      out.push({ text: line.text, x: sh.x + line.x, baseline: sh.y + line.baseline,
                 size: line.font_size, family: line.family, offs });
    }
  }
  return { width: pres.slide_width, height: pres.slide_height, lines: out };
}
"""


def pdf_chars(pdf: Path, page_no: int):
    """Every character of a page with its x and baseline, in points."""
    doc = pymupdf.open(pdf)
    page = doc[page_no - 1]
    out = []
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                for ch in span.get("chars", []):
                    out.append({"c": ch["c"], "x": ch["bbox"][0],
                                "y": span["origin"][1]})
    doc.close()
    return out


def match_line(line, chars):
    """The PDF characters that make up `line`: same text, one baseline, in order.

    Several lines can share a first character, so the candidate whose x is
    closest to the engine's wins -- the question being asked is about the
    ADVANCES, not about which line is which.
    """
    best = None
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
        if run and (best is None
                    or abs(run[0]["x"] - line["x"]) < abs(best[0]["x"] - line["x"])):
            best = run
    return best


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", default="d15")
    ap.add_argument("--slide", type=int, default=1)
    args = ap.parse_args()

    pptx = next(iter(sorted(
        (REPO / "pipeline_data/pptx_benchmark/dev/pptx").glob(args.deck + "*.pptx"))), None)
    pdf = next(iter(sorted(
        (REPO / "pipeline_data/pptx_benchmark/dev/pdf").glob(args.deck + "*.pdf"))), None)
    if not pptx or not pdf:
        sys.exit(f"no deck/pdf for {args.deck}")

    os.chdir(REPO / "web")
    httpd = socketserver.TCPServer(("127.0.0.1", 0),
                                   http.server.SimpleHTTPRequestHandler)
    httpd.RequestHandlerClass.log_message = lambda *a, **k: None
    port = httpd.server_address[1]
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    with sync_playwright() as p:
        b = p.chromium.launch()
        page = b.new_page()
        page.goto(f"http://127.0.0.1:{port}/pptx-editor.html")
        page.wait_for_function(
            "() => document.getElementById('status').textContent.includes('ready')")
        res = page.evaluate(JS, [list(pptx.read_bytes()), args.slide])
        b.close()

    chars = pdf_chars(pdf, args.slide)
    print(f"{args.deck} slide {args.slide}: {len(res['lines'])} engine lines, "
          f"{len(chars)} PDF characters\n")
    worst_all = []
    for line in res["lines"]:
        best = match_line(line, chars)
        if not best:
            continue
        # Compare each character's start, taking the PDF's first character as
        # the origin so a whole-line offset (alignment, insets) does not hide
        # the question being asked, which is about the ADVANCES.
        d = [abs((line["offs"][k] ) - (best[k]["x"] - best[0]["x"]))
             for k in range(len(best))]
        worst_all.append(max(d))
        flag = "  <-- " if max(d) > 0.5 else ""
        print(f"  {line['size']:6.2f}pt {line['family'][:18]:19} "
              f"n={len(best):3}  mean {sum(d)/len(d):.3f}pt  max {max(d):.3f}pt"
              f"{flag}{line['text'][:34]!r}")
    if worst_all:
        print(f"\n{len(worst_all)} lines matched -- worst character offset "
              f"{max(worst_all):.3f}pt, mean of worsts {sum(worst_all)/len(worst_all):.3f}pt")
    else:
        print("no line could be matched to the PDF")


if __name__ == "__main__":
    main()
