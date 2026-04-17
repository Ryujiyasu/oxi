"""Measure Word's per-row heights for d77a tables.

Uses COM Tables(i).Rows(j).Height and .HeightRule for each row.
Also measures the Y of the first cell's first line via cell.Range.Information(6)
to determine table top border.
"""
import os, sys, time, json
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

docx_path = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path, ReadOnly=True); time.sleep(0.5)
    n = doc.Tables.Count
    print(f"Tables: {n}")

    data = []
    for i in range(1, n + 1):
        t = doc.Tables(i)
        tbl_page = t.Range.Information(3)
        n_rows = t.Rows.Count
        n_cols = t.Columns.Count if hasattr(t, 'Columns') else 0
        # First cell first char Y
        try:
            first_cell = t.Cell(1, 1)
            first_y = first_cell.Range.Information(6)
        except Exception:
            first_y = None

        rows_data = []
        total_h_prop = 0.0
        for j in range(1, n_rows + 1):
            try:
                r = t.Rows(j)
                rh = r.Height  # in points
                rule = r.HeightRule  # wdRowHeightAtLeast=1, Exactly=2, Auto=0
                # Cell 1 Y for this row
                cell_y = r.Cells(1).Range.Information(6)
                rows_data.append({
                    "row": j, "h_prop": round(rh, 2), "rule": rule, "y": round(cell_y, 2)
                })
                total_h_prop += rh
            except Exception as e:
                rows_data.append({"row": j, "error": str(e)})

        # Get table bottom y via last row's next-line
        try:
            last_row = t.Rows(n_rows)
            last_cell = last_row.Cells(1)
            last_end = last_cell.Range.Information(6)
        except Exception:
            last_end = None

        entry = {
            "idx": i,
            "page": int(tbl_page) if tbl_page else None,
            "n_rows": n_rows,
            "n_cols": n_cols,
            "first_y": round(first_y, 2) if first_y else None,
            "rows": rows_data,
            "total_h_prop": round(total_h_prop, 2),
            "last_row_y": round(last_end, 2) if last_end else None,
        }
        data.append(entry)
        print(f"#{i} p{int(tbl_page)}: rows={n_rows} cols={n_cols} first_y={first_y:.1f} total_h={total_h_prop:.1f}")
        for r in rows_data[:3]:
            print(f"  row{r.get('row')}: y={r.get('y')} h={r.get('h_prop')} rule={r.get('rule')}")
        if len(rows_data) > 3:
            print(f"  ... ({len(rows_data)-3} more rows)")

    out = "pipeline_data/d77a_word_table_rows_v2.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nSaved: {out}")

    doc.Close(False)
finally:
    word.Quit()
