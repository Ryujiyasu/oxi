"""Measure Word's table row heights on d77a p.10 (top row-snap regression).

Goal: find a refined condition for when row grid-snap applies. If Word renders
some rows at content_h (no snap) and others at pitch multiple (snap), identify
the distinguishing property.
"""
import win32com.client, os, json

DOC = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    # Iterate tables; print rows on page 10 with measured heights
    for ti in range(1, wdoc.Tables.Count + 1):
        tbl = wdoc.Tables(ti)
        start_row_range = tbl.Rows(1).Range
        start_page = start_row_range.Information(3)
        end_row_range = tbl.Rows(tbl.Rows.Count).Range
        end_page = end_row_range.Information(3)
        # Skip if doesn't cover page 10
        if start_page > 10 or end_page < 10:
            continue
        first_cell = tbl.Cell(1,1).Range.Text.strip()[:30]
        print(f"\n=== Table {ti} (pages {start_page}-{end_page}): [{first_cell}]")
        positions = []
        for ri in range(1, tbl.Rows.Count + 1):
            try:
                rng = tbl.Rows(ri).Range
                pg = rng.Information(3)
                y = rng.Information(6)
                try: h = tbl.Rows(ri).Height
                except: h = None
                try: rule = tbl.Rows(ri).HeightRule
                except: rule = None
                positions.append((ri, pg, y, h, rule))
            except: pass
        # Report only rows on page 10
        for idx in range(len(positions) - 1):
            ri, pg, y, h, rule = positions[idx]
            if pg != 10: continue
            nri, npg, ny, nh, nrule = positions[idx+1]
            rh = round(ny - y, 2) if (y and ny and pg==npg) else None
            print(f"  row{ri:3d}  p{pg} y={y:.2f} set_h={h} rule={rule} rendered_h={rh}")
        # Last row if on p10
        if positions:
            ri, pg, y, h, rule = positions[-1]
            if pg == 10:
                print(f"  row{ri:3d}  p{pg} y={y:.2f} set_h={h} rule={rule} rendered_h=None (last)")
finally:
    wdoc.Close(False)
    word.Quit()
