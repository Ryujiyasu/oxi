# -*- coding: utf-8 -*-
"""Measure Spec #4 wave-3 truth: PDF spans + baselines for the FINE sweep.

Same as wave-2 (spec4b) but for spec4c_lspacing.pptx (4 slides, fine n sweep,
Calibri 18). PDF baseline deltas between consecutive paragraphs pin each
paragraph's line advance; the whole point is the fine f(n) shape + snap
detection + the sub-1.0 region.

Usage:
  python measure_pptx_spec4c.py <input.pptx> <out_dir>
"""
import json
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")

import fitz  # PyMuPDF

MSO_SHAPE_TYPE = {1: "AutoShape", 6: "Group", 14: "Placeholder", 17: "TextBox", 19: "Table"}


def _shape_label(t):
    return MSO_SHAPE_TYPE.get(int(t), "type_%d" % int(t))


def _paras_text(sh):
    try:
        if sh.HasTextFrame and sh.TextFrame.HasText:
            tr = sh.TextFrame.TextRange
            n = int(tr.Paragraphs().Count)
            out = []
            for i in range(1, n + 1):
                try:
                    out.append(str(tr.Paragraphs(i).Text))
                except Exception:
                    out.append(None)
            return out
    except Exception:
        pass
    return None


def _wrap(sh):
    try:
        if sh.HasTextFrame:
            return int(sh.TextFrame.WordWrap)
    except Exception:
        pass
    return None


def _textframe_margins(sh):
    try:
        tf = sh.TextFrame
        rec = {
            "margin_top": round(float(tf.MarginTop), 3),
            "margin_bottom": round(float(tf.MarginBottom), 3),
            "margin_left": round(float(tf.MarginLeft), 3),
            "margin_right": round(float(tf.MarginRight), 3),
        }
        try:
            rec["anchor"] = int(tf.VerticalAnchor)
        except Exception:
            rec["anchor"] = None
        return rec
    except Exception:
        return None


def measure_com(in_path):
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(os.path.abspath(in_path), True, False, False)
        time.sleep(0.5)
        n_slides = int(pres.Slides.Count)
        sw = round(float(pres.PageSetup.SlideWidth), 3)
        sh_h = round(float(pres.PageSetup.SlideHeight), 3)
        slides = []
        for si in range(1, n_slides + 1):
            slide = pres.Slides(si)
            shapes = []
            n_shapes = int(slide.Shapes.Count)
            for k in range(1, n_shapes + 1):
                s = slide.Shapes(k)
                rec = {
                    "idx": k,
                    "name": str(s.Name),
                    "type": _shape_label(s.Type),
                    "left": round(float(s.Left), 3),
                    "top": round(float(s.Top), 3),
                    "width": round(float(s.Width), 3),
                    "height": round(float(s.Height), 3),
                }
                rec["paras"] = _paras_text(s)
                rec["wrap"] = _wrap(s)
                rec["margins"] = _textframe_margins(s)
                shapes.append(rec)
            slides.append({"index": si, "slide_width": sw, "slide_height": sh_h, "shapes": shapes})
        return slides
    finally:
        try:
            pres.Close()
        except Exception:
            pass
        app.Quit()


def measure_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    pages = []
    for page in doc:
        rect = page.rect
        lines = []
        blocks = page.get_text("dict")["blocks"]
        for b in blocks:
            if b["type"] != 0:
                continue
            for ln in b["lines"]:
                max_origin_y = None
                x0 = None
                x1 = None
                txt = ""
                font = None
                size = None
                spans = []
                for sp in ln["spans"]:
                    oy = sp["origin"][1]
                    if max_origin_y is None or oy > max_origin_y:
                        max_origin_y = oy
                    sx0 = sp["bbox"][0]
                    sx1 = sp["bbox"][2]
                    if x0 is None or sx0 < x0:
                        x0 = sx0
                    if x1 is None or sx1 > x1:
                        x1 = sx1
                    t = sp["text"]
                    txt += t
                    if t.strip() and font is None:
                        font = sp.get("font")
                        size = round(float(sp.get("size", 0)), 3)
                    spans.append(
                        {
                            "text": t,
                            "font": sp.get("font"),
                            "size": round(float(sp.get("size", 0)), 3),
                            "origin": [round(sp["origin"][0], 3), round(sp["origin"][1], 3)],
                            "bbox": [round(v, 3) for v in sp["bbox"]],
                        }
                    )
                lines.append(
                    {
                        "y0": round(max_origin_y, 3),
                        "x0": round(x0, 3) if x0 is not None else None,
                        "x1": round(x1, 3) if x1 is not None else None,
                        "text": txt,
                        "font": font,
                        "size": size,
                        "spans": spans,
                    }
                )
        lines.sort(key=lambda l: (l["y0"], l["x0"] or 0))
        pages.append(
            {
                "page": page.number + 1,
                "width": round(float(rect.width), 3),
                "height": round(float(rect.height), 3),
                "lines": lines,
            }
        )
    doc.close()
    return pages


def main():
    in_path = sys.argv[1]
    out_dir = sys.argv[2]
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(in_path))[0]

    slides = measure_com(in_path)

    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(os.path.abspath(in_path), True, False, False)
        time.sleep(0.5)
        pdf_stem = os.path.join(out_dir, base)
        pres.SaveAs(pdf_stem, 32)
        pres.Close()
    finally:
        app.Quit()
    pdf_path = pdf_stem + ".pdf"
    if not os.path.exists(pdf_path):
        pdf_path = os.path.join(out_dir, "deck.pdf")
    pages = measure_pdf(pdf_path)

    with open(os.path.join(out_dir, "spec4c_truth.json"), "w", encoding="utf-8") as f:
        json.dump({"slides": slides}, f, ensure_ascii=False, indent=1)
    with open(os.path.join(out_dir, "pdf_lines.json"), "w", encoding="utf-8") as f:
        json.dump({"pages": pages}, f, ensure_ascii=False, indent=1)
    print("wrote", os.path.join(out_dir, "spec4c_truth.json"))
    print("wrote", os.path.join(out_dir, "pdf_lines.json"))
    print("wrote", pdf_path)


if __name__ == "__main__":
    main()
