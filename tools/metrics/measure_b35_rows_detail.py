"""Rich COM sweep of b35 Table 1 rows 2-5: per-cell paragraph structure.

For each row, enumerate each cell's paragraphs with:
- font, size, empty/content
- Y position of each paragraph's first line (within cell)
- Cell content height (approximation: last_y - first_y + leading)

Goal: derive multi-font cell formula. Previous attempts with 3 data points
(rows 1/3/5 only) failed; need cell-level detail to handle mixed-font cells.
"""
import json, time, sys
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path(r"C:/Users/ryuji/oxi-4/tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx")

word = win32com.client.Dispatch("Word.Application")
time.sleep(1.0)
word.Visible = False
word.DisplayAlerts = False

OUT = Path(__file__).with_name("output") / "b35_rows_detail.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

try:
    doc = word.Documents.Open(str(DOCX.resolve()), ReadOnly=True)
    time.sleep(1.5)

    # Table 1 is the main table on page 1.
    if doc.Tables.Count < 1:
        print("No tables"); sys.exit(1)

    tbl = doc.Tables(1)
    n_rows = tbl.Rows.Count
    print(f"Table 1: {n_rows} rows")

    rows_data = []
    for ri in range(1, min(n_rows + 1, 9)):  # rows 1-8
        row = tbl.Rows(ri)
        row_y = None
        try:
            row_y = row.Range.Information(6)
        except Exception:
            pass
        n_cells = row.Cells.Count
        row_info = {"row": ri, "row_y": round(row_y, 2) if row_y else None, "cells": []}
        print(f"\n=== Row {ri}: {n_cells} cells, y={row_y} ===")
        for ci in range(1, n_cells + 1):
            cell = row.Cells(ci)
            cell_range = cell.Range
            cell_x = cell_range.Information(5)
            cell_y = cell_range.Information(6)
            # Enumerate paragraphs
            paras = []
            for p in cell_range.Paragraphs:
                pf = p.Range.Font
                try:
                    py = p.Range.Information(6)
                    px = p.Range.Information(5)
                except Exception:
                    continue
                text = p.Range.Text.replace('\r','\\r').replace('\x07','\\a')[:40]
                empty = len(p.Range.Text.strip()) == 0 or p.Range.Text.strip() == '\x07'
                paras.append({
                    "y": round(py, 2),
                    "x": round(px, 2),
                    "fs": pf.Size,
                    "font": pf.Name,
                    "empty": empty,
                    "text": text,
                })
            cell_info = {"col": ci, "cell_x": round(cell_x, 2), "cell_y": round(cell_y, 2), "n_paras": len(paras), "paras": paras}
            row_info["cells"].append(cell_info)
            print(f"  Cell {ci}: x={cell_x:.1f} y={cell_y:.1f} paras={len(paras)}")
            for p in paras:
                e = '(e)' if p["empty"] else '   '
                print(f"    {e} y={p['y']:>7.2f} fs={p['fs']:>4} font={p['font'][:10]:<10} text={p['text']!r}")
        rows_data.append(row_info)

    doc.Close(False)

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(rows_data, f, ensure_ascii=False, indent=2)
    print(f"\nSaved → {OUT}")

finally:
    try: word.Quit()
    except: pass
