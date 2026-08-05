# -*- coding: utf-8 -*-
"""Measure PowerPoint truth for a .pptx via COM (the pptx analog of
measure_pagination_word.py).

For each slide it records:
  - slide size (points)
  - per-shape: index, name, mso type (+label), Left/Top/Width/Height (pt),
    Rotation, text (TextFrame.TextRange.Text), placeholder info, table dims

It also exports the deck to a PDF (ExportAsFixedFormat, ppFixedFormatTypePDF)
at the same absolute path minus extension -> the pixel/render truth for fitz.

Usage:
  python measure_pptx_word.py <input.pptx> <out_dir>
    -> <out_dir>/truth.json        (geometry + text)
    -> <out_dir>/deck.pdf          (PowerPoint's own render)
"""
import json
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")

# Empirically-confirmed Shape.Type values from the local PowerPoint 16.0
# (verified via _pptx_type_probe.py against semantic properties: HasTable /
# PlaceholderFormat.Type / HasTextFrame). Only CONFIRMED values are labeled;
# anything else falls back to "type_N" until it is probed and pinned.
#   1  = AutoShape            (RECTANGLE/OVAL/ROUNDED_RECTANGLE/CHEVRON ...)
#   14 = Placeholder          (Title -> PlaceholderFormat.Type 3, Subtitle -> 4)
#   17 = TextBox
#   19 = Table                (HasTable == True)
MSO_SHAPE_TYPE = {
    1: "AutoShape",
    6: "Group",
    14: "Placeholder",
    17: "TextBox",
    19: "Table",
}


def _shape_label(t):
    return MSO_SHAPE_TYPE.get(int(t), "type_%d" % int(t))


def _shape_text(sh):
    """Read the shape's text, truncated to 400 chars. Some shapes (tables)
    have TextFrame with no TextRange."""
    try:
        if sh.HasTextFrame and sh.TextFrame.HasText:
            return str(sh.TextFrame.TextRange.Text)[:400]
    except Exception:
        pass
    return None


def _table_dims(sh):
    # Empirically a table reports Type==19, but the robust discriminator is the
    # HasTable property (no magic shape-type number).
    try:
        if bool(sh.HasTable):
            return [int(sh.Table.Rows.Count), int(sh.Table.Columns.Count)]
    except Exception:
        pass
    return None


def _table_detail(sh):
    """Full table truth: per-column width, per-row height, per-cell text.
    Columns(i).Width / Rows(i).Height are in points; Cell(r,c) is 1-based."""
    try:
        if bool(sh.HasTable):
            tbl = sh.Table
            rows = int(tbl.Rows.Count)
            cols = int(tbl.Columns.Count)
            col_widths = []
            for c in range(1, cols + 1):
                try:
                    col_widths.append(round(float(tbl.Columns(c).Width), 3))
                except Exception:
                    col_widths.append(None)
            row_heights = []
            for r in range(1, rows + 1):
                try:
                    row_heights.append(round(float(tbl.Rows(r).Height), 3))
                except Exception:
                    row_heights.append(None)
            cells = []
            for r in range(1, rows + 1):
                row_cells = []
                for c in range(1, cols + 1):
                    try:
                        txt = str(tbl.Cell(r, c).Shape.TextFrame.TextRange.Text)
                    except Exception:
                        txt = None
                    row_cells.append(txt)
                cells.append(row_cells)
            return {
                "rows": rows,
                "cols": cols,
                "col_widths": col_widths,
                "row_heights": row_heights,
                "cells": cells,
            }
    except Exception:
        pass
    return None


def _placeholder_info(sh):
    try:
        pf = sh.PlaceholderFormat
        return {"type": int(pf.Type), "idx": int(pf.Idx)}
    except Exception:
        return None


def _shape_rec(sh, idx):
    rec = {
        "index": idx,
        "name": str(sh.Name),
        "type": int(sh.Type),
        "type_label": _shape_label(sh.Type),
    }
    # Geometry (points). Shape.Left/Top are relative to the slide's coordinate
    # origin, independent of zoom.
    try:
        rec["left"] = round(float(sh.Left), 3)
        rec["top"] = round(float(sh.Top), 3)
        rec["width"] = round(float(sh.Width), 3)
        rec["height"] = round(float(sh.Height), 3)
    except Exception:
        pass
    try:
        rec["rotation"] = round(float(sh.Rotation), 3)
    except Exception:
        pass
    text = _shape_text(sh)
    if text is not None:
        rec["text"] = text
    ph = _placeholder_info(sh)
    if ph:
        rec["placeholder"] = ph
    td = _table_dims(sh)
    if td:
        rec["table"] = td
    tdet = _table_detail(sh)
    if tdet:
        rec["table_detail"] = tdet
    # Group shapes: recurse children
    try:
        if int(sh.Type) == 6 and sh.GroupItems.Count > 0:
            children = []
            for i in range(1, int(sh.GroupItems.Count) + 1):
                try:
                    children.append(_shape_rec(sh.GroupItems(i), i - 1))
                except Exception:
                    pass
            rec["children"] = children
    except Exception:
        pass
    return rec


def measure_pptx(input_path, out_dir):
    import win32com.client
    import win32com.client.gencache as gencache

    os.makedirs(out_dir, exist_ok=True)
    abs_in = os.path.abspath(input_path)
    pdf_path = os.path.join(os.path.abspath(out_dir), "deck.pdf")

    # PowerPoint type library -> early (makepy) binding. Without this,
    # dynamic dispatch cannot resolve Presentation.SlideWidth / Shape.Rotation
    # etc. (AttributeError on late-bound property access).
    try:
        gencache.EnsureModule(
            "{91493440-5A91-11CF-8700-00AA0060263B}", 0, 1, 0
        )
    except Exception as e:
        print("gencache EnsureModule failed (falling back):", e)

    # DispatchEx = fresh PowerPoint instance (never attach to a running one).
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        # WithWindow=False so nothing flashes; Visible can stay False.
        pres = app.Presentations.Open(abs_in, True, False, False)
        try:
            # Slide dimensions live on PageSetup, not on Presentation directly.
            width = round(float(pres.PageSetup.SlideWidth), 3)
            height = round(float(pres.PageSetup.SlideHeight), 3)
            n_slides = int(pres.Slides.Count)
            slides = []
            for si in range(1, n_slides + 1):
                slide = pres.Slides(si)
                shapes = []
                n_shapes = int(slide.Shapes.Count)
                for i in range(1, n_shapes + 1):
                    try:
                        shapes.append(_shape_rec(slide.Shapes(i), i - 1))
                    except Exception as e:
                        shapes.append({"index": i - 1, "error": str(e)})
                slides.append({"index": si, "shapes": shapes})
            # Export the deck to PDF as render truth.
            # ExportAsFixedFormat fails under dynamic dispatch (win32com can't
            # coerce its args without the typelib); SaveAs with ppSaveAsPDF=32
            # is the classic alternative and works late-bound.
            pres.SaveAs(pdf_path, 32)
            # Give PowerPoint a moment to finish writing.
            for _ in range(20):
                if os.path.exists(pdf_path):
                    break
                time.sleep(0.1)
            result = {
                "input": abs_in,
                "slide_width": width,
                "slide_height": height,
                "slide_count": n_slides,
                "slides": slides,
                "pdf": pdf_path,
            }
        finally:
            try:
                pres.Close()
            except Exception:
                pass
    finally:
        try:
            app.Quit()
        except Exception:
            pass

    with open(os.path.join(out_dir, "truth.json"), "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)
    return result


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python measure_pptx_word.py <input.pptx> <out_dir>")
        sys.exit(1)
    r = measure_pptx(sys.argv[1], sys.argv[2])
    print("slides=%d size=%.1fx%.1f shapes=sum(%s)" % (
        r["slide_count"], r["slide_width"], r["slide_height"],
        [len(s["shapes"]) for s in r["slides"]]))
