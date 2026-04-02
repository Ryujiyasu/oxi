"""Measure table row heights via Word COM API.

Uses Cell-based approach to handle merged cells.
For each table, measures actual row Y positions from cell paragraphs.

Usage:
  python measure_table_row_height.py <docx_path>
"""
import win32com.client
import sys
import os
import json
from collections import defaultdict

def measure(docx_path):
    docx_path = os.path.abspath(docx_path)
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        results = []

        for ti in range(1, doc.Tables.Count + 1):
            table = doc.Tables(ti)
            table_data = {"table": ti, "rows": []}

            # Collect Y positions per row index via cells
            row_ys = defaultdict(list)
            row_info = {}

            for ci in range(1, table.Range.Cells.Count + 1):
                try:
                    cell = table.Range.Cells(ci)
                    ri = cell.RowIndex
                    y = cell.Range.Paragraphs(1).Range.Information(6)
                    row_ys[ri].append(y)

                    if ri not in row_info:
                        try:
                            row = cell.Row
                            height_rule_map = {0: "auto", 1: "atLeast", 2: "exact"}
                            row_info[ri] = {
                                "height": round(row.Height, 2),
                                "height_rule": height_rule_map.get(row.HeightRule, str(row.HeightRule)),
                            }
                            # Border widths
                            try:
                                row_info[ri]["border_top"] = round(row.Borders(-1).LineWidth / 8.0, 2)
                                row_info[ri]["border_bottom"] = round(row.Borders(-3).LineWidth / 8.0, 2)
                            except:
                                pass
                        except:
                            row_info[ri] = {}
                except:
                    continue

            # Build sorted row data
            sorted_rows = sorted(row_ys.keys())
            for i, ri in enumerate(sorted_rows):
                y_min = min(row_ys[ri])
                info = row_info.get(ri, {})
                rd = {
                    "row": ri,
                    "y": round(y_min, 2),
                    **info,
                }
                # Compute actual height from next row
                if i + 1 < len(sorted_rows):
                    next_y = min(row_ys[sorted_rows[i + 1]])
                    rd["actual_height"] = round(next_y - y_min, 2)
                table_data["rows"].append(rd)

            results.append(table_data)

        doc.Close(False)
        return results
    finally:
        word.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python measure_table_row_height.py <docx_path>")
        sys.exit(1)

    results = measure(sys.argv[1])
    print(json.dumps(results, indent=2, ensure_ascii=False))
