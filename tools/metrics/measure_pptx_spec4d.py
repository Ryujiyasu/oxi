# -*- coding: utf-8 -*-
"""Spec #4 wave-4 measure: read the spec4d_multiline.pptx back via COM and
export to PDF; read per-line baselines from the PDF.

Per slide we print:
  - n (line_spacing as reported by COM / the known sweep value)
  - the per-line baselines (y0)
  - the consecutive baseline deltas WITHIN the single multi-line paragraph
    (excluding the first line whose delta mixes the frame-top baseline offset)

The within-paragraph deltas are the clean line advance for that n.

Outputs under pipeline_data\\pptx_probes\\spec4d_truth\\:
  spec4d_truth.json  (COM: line_spacing per paragraph, margins, wrap, anchor)
  pdf_lines.json     (PDF: per-line y0/x0/x1/text/font/size)
  spec4d_multiline.pdf
"""
import json
import os

import win32com.client

from pptx import Presentation

import sys
sys.stdout.reconfigure(encoding="utf-8")

BASE = os.path.abspath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "..",
    "pipeline_data", "pptx_probes"))
SRC = os.path.join(BASE, "spec4d_multiline.pptx")
PDF_STEM = os.path.join(BASE, "spec4d_multiline")
OUTDIR = os.path.join(BASE, "spec4d_truth")


def measure_com():
    pres = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        p = pres.Presentations.Open(SRC, ReadOnly=True, WithWindow=False)
        rows = []
        for s in p.Slides:
            for sh in s.Shapes:
                if sh.HasTextFrame:
                    tf = sh.TextFrame
                    tr = tf.TextRange
                    pf = tr.Paragraphs(1).ParagraphFormat
                    def g(a):
                        try:
                            return round(float(getattr(pf, a)), 3)
                        except Exception:
                            return None
                    rows.append({
                        "slide": s.SlideIndex,
                        "rule_within": g("LineRuleWithin"),
                        "line_spacing": g("LineSpacing"),
                        "space_before": g("SpaceBefore"),
                        "space_after": g("SpaceAfter"),
                        "wrap": tf.WordWrap,
                        "anchor": tf.VerticalAnchor,
                        "margin_top": tf.MarginTop,
                        "margin_bottom": tf.MarginBottom,
                        "margin_left": tf.MarginLeft,
                        "margin_right": tf.MarginRight,
                        "font_name": tr.Font.Name,
                        "font_size": tr.Font.Size,
                    })
        p.SaveAs(PDF_STEM, 32)  # 32 = ppSaveAsPDF
        p.Close()
        return rows
    finally:
        pres.Quit()


def measure_pdf(pdf_path):
    import fitz  # PyMuPDF
    doc = fitz.open(pdf_path)
    pages = []
    for pg in doc:
        lines = []
        for blk in pg.get_text("dict", sort=True)["blocks"]:
            for line in blk.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                origins = [sp["origin"] for sp in spans]
                y0 = max(o[1] for o in origins)
                x0 = min(sp["bbox"][0] for sp in spans)
                x1 = max(sp["bbox"][2] for sp in spans)
                font = spans[0]["font"]
                size = spans[0]["size"]
                text = "".join(sp["text"] for sp in spans)
                lines.append({
                    "y0": round(y0, 3), "x0": round(x0, 3), "x1": round(x1, 3),
                    "text": text, "font": font, "size": round(size, 4),
                    "spans": [{"origin": list(sp["origin"]), "bbox": list(sp["bbox"])}
                              for sp in spans],
                })
        lines.sort(key=lambda l: (l["y0"], l["x0"]))
        pages.append({"page": pg.number + 1, "width": pg.rect.width,
                      "height": pg.rect.height, "lines": lines})
    doc.close()
    return pages


def main():
    os.makedirs(OUTDIR, exist_ok=True)
    com = measure_com()
    pdf = measure_pdf(PDF_STEM + ".pdf")
    with open(os.path.join(OUTDIR, "spec4d_truth.json"), "w", encoding="utf-8") as f:
        json.dump(com, f, ensure_ascii=False, indent=1)
    with open(os.path.join(OUTDIR, "pdf_lines.json"), "w", encoding="utf-8") as f:
        json.dump(pdf, f, ensure_ascii=False, indent=1)
    print("wrote spec4d_truth.json / pdf_lines.json / spec4d_multiline.pdf")

    # quick per-slide readout
    print("\n=== per-slide line advances (within-paragraph deltas) ===")
    n_by_slide = {}
    for c in com:
        n_by_slide[c["slide"]] = c["line_spacing"]
    for p in pdf:
        n = n_by_slide.get(p["page"])
        ys = [l["y0"] for l in p["lines"]]
        if len(ys) >= 2:
            deltas = [round(ys[i + 1] - ys[i], 3) for i in range(len(ys) - 1)]
            print("slide%d n=%s lines=%d baselines=%s deltas=%s" % (
                p["page"], n, len(ys), ys, deltas))
        else:
            print("slide%d n=%s lines=%d baselines=%s" % (p["page"], n, len(ys), ys))


if __name__ == "__main__":
    main()
