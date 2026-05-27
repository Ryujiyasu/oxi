"""
S349 - Measure Word's actual line count for paragraphs in 31420af label cells.

Hypothesis: Word's lineRule="exact" + trHeight (auto) combination renders
content within row boundary, possibly clipping or fitting in fewer lines than
Oxi computes.

For each paragraph in problem cells (rows 2, 3, 8, 10 from XML), measure:
- Font name / size
- Character-by-character x/y positions to detect wrap
- Total line count
- Computed line height vs measured advance
"""
import os, json, sys
import win32com.client
from pathlib import Path

DOC_PATH = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\31420af1a08f_tokumei_08_07.docx").resolve()
OUT_PATH = Path(r"c:\Users\ryuji\oxi-main\tools\metrics\31420af_wrap_lines_word.json").resolve()

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = False

try:
    doc = word.Documents.Open(str(DOC_PATH), ReadOnly=True)
    doc.ActiveWindow.View.Type = 3
    doc.Repaginate()

    # Walk each table → each cell → for cells of interest (label paras)
    # measure per-char positions
    out = {"doc": str(DOC_PATH), "cells": []}

    tbl = doc.Tables(1)
    cells_coll = tbl.Range.Cells
    n_cells = cells_coll.Count
    target_seqs = [5, 7, 9, 11, 13]  # cells with row labels (1, 2, 3, 4, 5)
    # Also: cell containing content (e.g., bullet content in row 9 col 3)
    # That's cell with text "利用者の範囲は適正か" — should be in row 9 col 3
    # From earlier dump that was seq 18 (RowIdx=9, ColIdx=3)
    target_seqs += [18, 22, 24, 26]  # bullet-content cells

    for c_idx in range(1, n_cells + 1):
        if c_idx not in target_seqs:
            continue
        cell = cells_coll(c_idx)
        cell_info = {
            "cell_seq": c_idx,
            "row_index": getattr(cell, "RowIndex", None),
            "column_index": getattr(cell, "ColumnIndex", None),
            "width_pt": cell.Width,
            "height_pt": cell.Height,
            "paragraphs": [],
        }

        for p_i in range(1, cell.Range.Paragraphs.Count + 1):
            try:
                p = cell.Range.Paragraphs(p_i)
                p_rng = p.Range
                text = p_rng.Text.rstrip("\r\x07")
                # Per-char positions
                chars = []
                for cpos in range(p_rng.Start, p_rng.End):
                    cr = doc.Range(cpos, cpos + 1)
                    try:
                        ch_text = cr.Text
                        if not ch_text or ch_text in ("\r", "\x07"):
                            continue
                        y = cr.Information(6)
                        x = cr.Information(7)
                        chars.append({"ch": ch_text, "x": x, "y": y})
                    except Exception:
                        pass
                # Get font of first run
                fname = None
                fsize = None
                try:
                    fname = p_rng.Font.NameAscii if hasattr(p_rng.Font, "NameAscii") else p_rng.Font.Name
                    fname_far_east = getattr(p_rng.Font, "NameFarEast", None)
                    fsize = p_rng.Font.Size
                except Exception:
                    pass
                # Count distinct y values = visual line count
                ys = sorted(set(round(c["y"], 1) for c in chars))
                cell_info["paragraphs"].append({
                    "p_idx": p_i,
                    "text": text[:60],
                    "len": len(text),
                    "font_name": fname,
                    "font_name_far_east": fname_far_east if 'fname_far_east' in dir() else None,
                    "font_size_pt": fsize,
                    "n_chars_measured": len(chars),
                    "visible_y_lines": len(ys),
                    "y_values": ys,
                    "first_char": chars[0] if chars else None,
                    "last_char": chars[-1] if chars else None,
                })
            except Exception as e:
                cell_info["paragraphs"].append({"p_idx": p_i, "err": str(e)})

        out["cells"].append(cell_info)

    OUT_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote: {OUT_PATH}")
    print(f"Cells: {len(out['cells'])}")
finally:
    try:
        doc.Close(SaveChanges=False)
    except Exception:
        pass
    word.Quit()
