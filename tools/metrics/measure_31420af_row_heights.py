"""
S349 - Measure Word's actual rendered table for 31420af (vMerge-aware).

Word COM's tbl.Rows(N) errors with vertical merge present. Use tbl.Range.Cells
to iterate cells directly. Each cell exposes RowIndex/ColumnIndex.

For each cell, measure: x, y, width, height (Word's effective rendered cell),
v_align, n_paragraphs, first/last paragraph y.
"""
import os, json, sys
import win32com.client
from pathlib import Path

DOC_PATH = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\31420af1a08f_tokumei_08_07.docx").resolve()
OUT_PATH = Path(r"c:\Users\ryuji\oxi-main\tools\metrics\31420af_row_heights_word.json").resolve()

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = False

try:
    doc = word.Documents.Open(str(DOC_PATH), ReadOnly=True)
    doc.ActiveWindow.View.Type = 3  # wdPrintView
    doc.Repaginate()

    out = {"doc": str(DOC_PATH), "tables": []}

    for tbl_i, tbl in enumerate(doc.Tables, start=1):
        tbl_info = {"tbl_idx": tbl_i, "cells": []}

        cells_coll = tbl.Range.Cells
        n_cells = cells_coll.Count
        for c_idx in range(1, n_cells + 1):
            cell = cells_coll(c_idx)
            cell_info = {
                "cell_seq": c_idx,
                "row_index": getattr(cell, "RowIndex", None),
                "column_index": getattr(cell, "ColumnIndex", None),
                "width_pt": cell.Width if hasattr(cell, "Width") else None,
                "height_pt": cell.Height if hasattr(cell, "Height") else None,
                "v_align": cell.VerticalAlignment if hasattr(cell, "VerticalAlignment") else None,
            }
            try:
                rng = cell.Range
                start_rng = doc.Range(rng.Start, rng.Start)
                end_rng = doc.Range(rng.End, rng.End)
                cell_info["start_y_pt"] = start_rng.Information(6)
                cell_info["end_y_pt"] = end_rng.Information(6)
                cell_info["start_x_pt"] = start_rng.Information(7)
                cell_info["start_page"] = start_rng.Information(3)
                cell_info["end_page"] = end_rng.Information(3)
            except Exception as e:
                cell_info["range_err"] = str(e)

            try:
                paras = []
                for p_i in range(1, cell.Range.Paragraphs.Count + 1):
                    p = cell.Range.Paragraphs(p_i)
                    p_rng = p.Range
                    try:
                        p_start = doc.Range(p_rng.Start, p_rng.Start)
                        y = p_start.Information(6)
                        x = p_start.Information(7)
                        page = p_start.Information(3)
                    except Exception:
                        y = x = page = None
                    text = (p_rng.Text or "").rstrip("\r\x07")
                    paras.append({
                        "p_idx": p_i,
                        "y_pt": y,
                        "x_pt": x,
                        "page": page,
                        "text": text[:80],
                        "len": len(text),
                    })
                cell_info["paragraphs"] = paras
            except Exception as e:
                cell_info["para_err"] = str(e)
            tbl_info["cells"].append(cell_info)

        out["tables"].append(tbl_info)

    OUT_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote: {OUT_PATH}")
    if out["tables"]:
        n = len(out["tables"][0]["cells"])
        print(f"Tables: {len(out['tables'])}, cells in table 0: {n}")
finally:
    try:
        doc.Close(SaveChanges=False)
    except Exception:
        pass
    word.Quit()
