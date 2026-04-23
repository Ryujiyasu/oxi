"""Measure x positions of Word d77a p.7 paragraphs idx=109-112 to determine
if the table is 2-column (cells side-by-side) or 1-column (paragraphs stacked).

If idx=109 and idx=111 have same y but different x → 2-column layout.
If 1-column, Oxi's stacked interpretation is correct but line-count differs.
"""
import json
from pathlib import Path
import win32com.client as w32


DOC = Path(r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx")
OUT = Path(r"c:\Users\ryuji\oxi-main\pipeline_data\d77a_p7_table_structure.json")


def main():
    word = w32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    result = {"file": DOC.name}
    try:
        doc = word.Documents.Open(str(DOC.resolve()), ReadOnly=True)
        try:
            # Iterate idx=108..114 and capture (x, y, in_table, text)
            # Information(1) = wdHorizontalPositionRelativeToPage
            paras_info = []
            for i, p in enumerate(doc.Paragraphs, start=1):
                if i < 107 or i > 116:
                    continue
                r = p.Range
                try:
                    y = r.Information(6)
                    x = r.Information(1)  # wdHorizontalPositionRelativeToPage
                    pg = r.Information(3)
                    in_table = bool(r.Information(12))
                except Exception:
                    continue
                text = r.Text[:80].replace("\r", "\\r").replace("\x07", "\\x07").replace("\n", "\\n")
                paras_info.append({
                    "idx": i,
                    "page": pg,
                    "x_pt": round(x, 3),
                    "y_pt": round(y, 3),
                    "in_table": in_table,
                    "text": text,
                })
            result["paras"] = paras_info

            # Also enumerate tables on p.7 and their cell structure
            tables_on_p7 = []
            for ti, tbl in enumerate(doc.Tables, start=1):
                try:
                    y = tbl.Range.Information(6)
                    pg = tbl.Range.Information(3)
                except Exception:
                    continue
                if pg == 7:
                    rows = tbl.Rows.Count
                    cols = tbl.Columns.Count
                    # Per-cell positions
                    cell_info = []
                    for ri in range(1, rows + 1):
                        for ci in range(1, cols + 1):
                            try:
                                cell = tbl.Cell(ri, ci)
                                cx = cell.Range.Information(1)
                                cy = cell.Range.Information(6)
                                ctxt = cell.Range.Text[:40].replace("\r", "\\r").replace("\x07", "\\x07")
                                cell_info.append({
                                    "row": ri, "col": ci,
                                    "x_pt": round(cx, 2),
                                    "y_pt": round(cy, 2),
                                    "text": ctxt,
                                })
                            except Exception as e:
                                cell_info.append({"row": ri, "col": ci, "error": str(e)})
                    tables_on_p7.append({
                        "table_idx": ti,
                        "rows": rows,
                        "cols": cols,
                        "y_pt": round(y, 2),
                        "cells": cell_info,
                    })
            result["tables_on_p7"] = tables_on_p7
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    print("=== Paragraphs idx=107-116 ===")
    for p in result["paras"]:
        marker = "[TBL]" if p["in_table"] else "[BDY]"
        print(f"  idx={p['idx']:3d} p{p['page']} x={p['x_pt']:6.2f} y={p['y_pt']:6.2f} {marker} {p['text'][:50]!r}")

    print("\n=== Tables on p.7 ===")
    for t in result["tables_on_p7"]:
        print(f"  Table #{t['table_idx']}: {t['rows']} rows × {t['cols']} cols, top_y={t['y_pt']}")
        for c in t["cells"]:
            if "error" in c:
                print(f"    cell[{c['row']},{c['col']}]: ERROR {c['error']}")
            else:
                print(f"    cell[{c['row']},{c['col']}] x={c['x_pt']:6.2f} y={c['y_pt']:6.2f} text={c['text'][:40]!r}")


if __name__ == "__main__":
    main()
