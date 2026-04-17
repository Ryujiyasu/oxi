"""Measure Word's table row heights for 04b88e7e0b25_index-19.

Hypothesis: Oxi compresses table rows compared to Word, fitting more rows per
page → loses last page. Quantify by comparing Word's row heights (via COM) to
Oxi's (via layout_json).

Word API: Table.Rows(i).Height returns the rendered row height in points.
"""
import win32com.client, os, json

DOC = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\04b88e7e0b25_index-19.docx"

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    n_tables = wdoc.Tables.Count
    print(f"Document has {n_tables} table(s)")

    table_data = []
    for ti in range(1, n_tables + 1):
        tbl = wdoc.Tables(ti)
        first_row_range = tbl.Rows(1).Range
        page = first_row_range.Information(3)
        y_start = first_row_range.Information(6)
        first_cell_text = tbl.Cell(1, 1).Range.Text.strip()[:30] if tbl.Columns.Count else ""
        n_rows = tbl.Rows.Count
        n_cols = tbl.Columns.Count
        row_heights = []
        row_positions = []
        for ri in range(1, n_rows + 1):
            try:
                rng = tbl.Rows(ri).Range
                y = rng.Information(6)
                pg = rng.Information(3)
                # Height may raise if row has no explicit height
                try:
                    h = tbl.Rows(ri).Height
                except Exception:
                    h = None
                # HeightRule: 0=auto, 1=atLeast, 2=exact
                try:
                    rule = tbl.Rows(ri).HeightRule
                except Exception:
                    rule = None
                # Measure actual rendered height via next-row Y diff
                row_positions.append((ri, pg, y, h, rule))
            except Exception as e:
                row_positions.append((ri, None, None, None, None))
        # Compute rendered heights via consecutive Y positions where page matches
        rendered = []
        for idx in range(len(row_positions) - 1):
            ri, pg, y, h, rule = row_positions[idx]
            nri, npg, ny, nh, nrule = row_positions[idx+1]
            rh = None
            if y is not None and ny is not None and pg == npg:
                rh = round(ny - y, 2)
            rendered.append((ri, pg, round(y,2) if y is not None else None, h, rule, rh))
        # Last row rendered: use final row's position; can't compute from next
        ri, pg, y, h, rule = row_positions[-1]
        rendered.append((ri, pg, round(y,2) if y is not None else None, h, rule, None))

        # Print summary
        print(f"\n=== Table {ti}: starts p{page} y={y_start:.1f} - {n_rows}R x {n_cols}C - first cell: [{first_cell_text}]")
        # Print first 10 rows
        for ri, pg, y, h, rule, rh in rendered[:10]:
            hstr = f"{h:.2f}" if h else "None"
            print(f"  row{ri:3d}  page={pg}  y={y}  set_h={hstr}  rule={rule}  rendered_h={rh}")
        if len(rendered) > 15:
            print(f"  ... ({len(rendered)-15} middle rows)")
            for ri, pg, y, h, rule, rh in rendered[-5:]:
                hstr = f"{h:.2f}" if h else "None"
                print(f"  row{ri:3d}  page={pg}  y={y}  set_h={hstr}  rule={rule}  rendered_h={rh}")
        elif len(rendered) > 10:
            for ri, pg, y, h, rule, rh in rendered[10:]:
                hstr = f"{h:.2f}" if h else "None"
                print(f"  row{ri:3d}  page={pg}  y={y}  set_h={hstr}  rule={rule}  rendered_h={rh}")

        # Stats on rendered heights
        rh_vals = [rh for _,_,_,_,_,rh in rendered if rh is not None and rh > 0]
        if rh_vals:
            print(f"  rendered heights: min={min(rh_vals):.2f}, median={sorted(rh_vals)[len(rh_vals)//2]:.2f}, max={max(rh_vals):.2f}, n={len(rh_vals)}")
        table_data.append({
            "index": ti,
            "start_page": page,
            "start_y": y_start,
            "rows": rendered,
            "first_cell": first_cell_text,
            "n_rows": n_rows,
            "n_cols": n_cols,
        })

    # Save
    out = os.path.join(os.path.dirname(__file__), 'output', 'measure_04b88_table_rows.json')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(table_data, f, indent=2, ensure_ascii=False, default=str)
    print(f"\nSaved to {out}")
finally:
    wdoc.Close(False)
    word.Quit()
