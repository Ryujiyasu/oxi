# -*- coding: utf-8 -*-
"""Measure Spec #4 (text frame) truth: COM structure + PDF baseline geometry.

COM (PowerPoint 16.0) supplies structure per shape: type, geometry, paragraph
texts, wrap setting, line/paragraph spacing. The actual LINE ADVANCE and WRAP
positions are read from PowerPoint's own PDF export with fitz, because COM has
no per-line Information(6) analog.

For each slide page the PDF reader returns, per line: baseline y (the max
origin y among the line's spans = the baseline), min/max x, and the joined
text. Comparing consecutive baselines inside one shape pins the line advance.

Usage:
  python measure_pptx_spec4.py <input.pptx> <out_dir>
    -> <out_dir>/spec4_truth.json  (COM structure)
    -> <out_dir>/deck.pdf          (PowerPoint render)
    -> <out_dir>/pdf_lines.json    (fitz per-line baselines per page)
"""
import json
import os
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")

import fitz  # PyMuPDF

MSO_SHAPE_TYPE = {1: "AutoShape", 6: "Group", 14: "Placeholder", 17: "TextBox", 19: "Table"}


def _shape_label(t):
    return MSO_SHAPE_TYPE.get(int(t), "type_%d" % int(t))


def _paras_text(sh):
    """Per-paragraph texts (paragraph i = 1-based). Returns None if no text."""
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
            return int(sh.TextFrame.WordWrap)  # -1 msoTrue / 0 msoFalse
    except Exception:
        pass
    return None


def _line_spacing(sh):
    """Per-paragraph (LineRuleWithin, LineSpacing, LineRuleBefore, SpaceBefore,
    LineRuleAfter, SpaceAfter). Returns a list of dicts or None."""
    try:
        if not (sh.HasTextFrame and sh.TextFrame.HasText):
            return None
        tr = sh.TextFrame.TextRange
        n = int(tr.Paragraphs().Count)
        out = []
        for i in range(1, n + 1):
            pf = tr.Paragraphs(i).ParagraphFormat
            rec = {}
            for key, attr in (
                ("rule_within", "LineRuleWithin"),
                ("line_spacing", "LineSpacing"),
                ("space_before", "SpaceBefore"),
                ("space_after", "SpaceAfter"),
            ):
                try:
                    rec[key] = round(float(getattr(pf, attr)), 3)
                except Exception:
                    rec[key] = None
            out.append(rec)
        return out
    except Exception:
        return None


def measure_com(in_path):
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(os.path.abspath(in_path), True, False, False)
        time.sleep(0.5)
        n_slides = int(pres.Slides.Count)
        # Slide size is presentation-scoped (there is no per-slide PageSetup).
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
                rec["line_spacing"] = _line_spacing(s)
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
    """Return per-page: {page, width, height, lines:[{y0, x0, x1, text}]}
    where y0 is the max span-origin y (= baseline), x0/x1 the line extent."""
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
                    txt += sp["text"]
                lines.append(
                    {
                        "y0": round(max_origin_y, 3),
                        "x0": round(x0, 3) if x0 is not None else None,
                        "x1": round(x1, 3) if x1 is not None else None,
                        "text": txt,
                    }
                )
        # sort by y0 then x0
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

    pdf_path = os.path.join(out_dir, "deck.pdf")
    # SaveAs(ppFixedFormatTypePDF=32) -> absolute path minus extension.
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
    # PowerPoint appends .pdf
    if not os.path.exists(pdf_path):
        pdf_path = pdf_stem + ".pdf"
    pages = measure_pdf(pdf_path)

    with open(os.path.join(out_dir, "spec4_truth.json"), "w", encoding="utf-8") as f:
        json.dump({"slides": slides}, f, ensure_ascii=False, indent=1)
    with open(os.path.join(out_dir, "pdf_lines.json"), "w", encoding="utf-8") as f:
        json.dump({"pages": pages}, f, ensure_ascii=False, indent=1)
    print("wrote", os.path.join(out_dir, "spec4_truth.json"))
    print("wrote", os.path.join(out_dir, "pdf_lines.json"))
    print("wrote", pdf_path)


if __name__ == "__main__":
    main()
