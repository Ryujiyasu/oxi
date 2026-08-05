# -*- coding: utf-8 -*-
"""Spec #4e measure: PowerPoint COM + PDF for spec4e_multifont.pptx.

Reads per-slide ParagraphFormat via getattr (property access), exports PDF,
and runs measure_pdf to collect line baselines / ink bboxes.
"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

OUT_DIR = r"c:\Users\ryuji\oxi-main\pipeline_data\pptx_probes\spec4e_multifont"
OUT_PPTX = os.path.join(OUT_DIR, "spec4e_multifont.pptx")
PDF_STEM = os.path.join(OUT_DIR, "spec4e_multifont")
TRUTH_JSON = os.path.join(OUT_DIR, "spec4e_truth.json")
PDF_LINES_JSON = os.path.join(OUT_DIR, "pdf_lines.json")


def measure_pdf(pdf_path, out_json):
    import fitz
    doc = fitz.open(pdf_path)
    pages = []
    for pg in doc:
        d = pg.get_text("dict", sort=True)
        lines = []
        for blk in d.get("blocks", []):
            for ln in blk.get("lines", []):
                spans = []
                for sp in ln.get("spans", []):
                    x0, y0, x1, y1 = sp["bbox"]
                    spans.append({
                        "origin": [sp["origin"][0], sp["origin"][1]],
                        "bbox": [x0, y0, x1, y1],
                        "text": sp["text"],
                        "font": sp["font"],
                        "size": sp["size"],
                    })
                if not spans:
                    continue
                lines.append({
                    "y0": min(s["bbox"][1] for s in spans),
                    "x0": min(s["bbox"][0] for s in spans),
                    "x1": max(s["bbox"][2] for s in spans),
                    "text": "".join(s["text"] for s in spans),
                    "font": spans[0]["font"],
                    "size": spans[0]["size"],
                    "spans": spans,
                })
        pages.append({
            "page": len(pages) + 1,
            "width": pg.rect.width,
            "height": pg.rect.height,
            "lines": lines,
        })
    import json
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(pages, f, ensure_ascii=False)
    print("wrote %s (%d pages)" % (out_json, len(pages)))


def main():
    import json
    import win32com.client as win32

    truth = []
    app = win32.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(OUT_PPTX, WithWindow=False)
        sw = pres.PageSetup.SlideWidth
        sh = pres.PageSetup.SlideHeight
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            rec = {
                "slide_index": i,
                "slide_width_pt": sw / 72.0,
                "slide_height_pt": sh / 72.0,
                "shapes": [],
            }
            for si in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes(si)
                tf = shp.TextFrame
                tr = tf.TextRange
                para0 = tr.Paragraphs(1)
                pf = para0.ParagraphFormat
                shp_rec = {
                    "shape_index": si,
                    "left_pt": shp.Left / 72.0,
                    "top_pt": shp.Top / 72.0,
                    "width_pt": shp.Width / 72.0,
                    "height_pt": shp.Height / 72.0,
                    "word_wrap": bool(tf.WordWrap),
                    "vertical_anchor": int(tf.VerticalAnchor),
                    "margin_top_pt": tf.MarginTop / 72.0,
                    "margin_bottom_pt": tf.MarginBottom / 72.0,
                    "margin_left_pt": tf.MarginLeft / 72.0,
                    "margin_right_pt": tf.MarginRight / 72.0,
                    "rule_within": str(getattr(pf, "LineRuleWithin", None)),
                    "line_spacing": str(getattr(pf, "LineSpacing", None)),
                    "space_before": str(getattr(pf, "SpaceBefore", None)),
                    "space_after": str(getattr(pf, "SpaceAfter", None)),
                    "font_name": str(tr.Font.Name),
                    "font_size": str(tr.Font.Size),
                }
                rec["shapes"].append(shp_rec)
            truth.append(rec)
        pres.SaveAs(PDF_STEM, 32)
        pres.Close()
    finally:
        app.Quit()

    with open(TRUTH_JSON, "w", encoding="utf-8") as f:
        json.dump(truth, f, ensure_ascii=False, indent=1)
    print("wrote %s" % TRUTH_JSON)

    measure_pdf(PDF_STEM + ".pdf", PDF_LINES_JSON)


if __name__ == "__main__":
    main()
