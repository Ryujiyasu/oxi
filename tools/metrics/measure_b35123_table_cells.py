"""Cluster A verification: b35123 table cell internal Y position.

b35123_tokumei_08_01.docx has 22 tables in 78 paragraphs (28% table density).
Min SSIM 0.666 at page 1.

Measurement strategy:
- Open in Word, iterate Tables(i)
- For each cell: measure cell.Range.Information(6) for top, cell.Cell.Range
  per-paragraph y values
- Compare to Oxi cached layout per-cell text element y values
"""
import json
import sys
import time
from pathlib import Path
import pythoncom
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path("tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx")
OUT = Path("pipeline_data/b35123_table_cells_measurement.json")


def measure(word, doc_path):
    last = None
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(doc_path.resolve()), ReadOnly=True)
            time.sleep(0.5)
            n_tables = doc.Tables.Count
            print(f"  Found {n_tables} tables")
            results = []
            for ti in range(1, n_tables + 1):
                tbl = doc.Tables(ti)
                rows = tbl.Rows.Count
                cols = tbl.Columns.Count
                tbl_top_y = tbl.Range.Information(6)
                tbl_page = tbl.Range.Information(3)
                tbl_data = {
                    "table_idx": ti,
                    "tbl_top_y": round(tbl_top_y, 2),
                    "tbl_page": tbl_page,
                    "n_rows": rows,
                    "n_cols": cols,
                    "rows": [],
                }
                # Sample first 3 rows × all cols + last row
                row_indices = sorted(set(list(range(1, min(rows + 1, 4))) + ([rows] if rows > 3 else [])))
                for ri in row_indices:
                    row_data = {"row_idx": ri, "cells": []}
                    try:
                        ncells_in_row = tbl.Rows(ri).Cells.Count
                    except Exception:
                        continue
                    for ci in range(1, min(ncells_in_row + 1, 4)):  # first 3 cols only
                        try:
                            cell = tbl.Cell(ri, ci)
                            cell_top_y = cell.Range.Information(6)
                            cell_x = cell.Range.Information(5)
                            # Get first paragraph's first char y
                            cr = cell.Range
                            first_char_y = cr.Information(6)
                            text = (cr.Text or "")[:30].replace("\r","\\r").replace("\x07","\\x07")
                            row_data["cells"].append({
                                "col_idx": ci,
                                "cell_top_y": round(cell_top_y, 2),
                                "cell_x": round(cell_x, 2),
                                "first_char_y": round(first_char_y, 2),
                                "text": text,
                            })
                        except Exception:
                            continue
                    tbl_data["rows"].append(row_data)
                results.append(tbl_data)
                if ti <= 3:
                    print(f"  Table {ti} (page {tbl_page}): {rows}r × {cols}c, top_y={round(tbl_top_y,2)}")
            doc.Close(SaveChanges=False)
            return results
        except Exception as e:
            last = e
            time.sleep(0.8 + attempt * 0.5)
    return [{"error": str(last)}]


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    try:
        results = measure(word, DOCX)
    finally:
        try: word.Quit()
        except: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
