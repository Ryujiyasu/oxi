"""Simpler: use the first cell in each row as the Y anchor."""
import win32com.client, os

DOC = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx"
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    n_tables = wdoc.Tables.Count
    print(f"Document has {n_tables} table(s)")
    for ti in range(1, n_tables + 1):
        tbl = wdoc.Tables(ti)
        n_rows = tbl.Rows.Count
        first_cell = tbl.Cell(1,1).Range.Text.strip()[:30]
        print(f"\n=== Table {ti} ({n_rows} rows): [{first_cell}]")
        ys = []
        for ri in range(1, n_rows + 1):
            try:
                # Use first cell of each row
                c = tbl.Cell(ri, 1)
                rng = c.Range
                pg = rng.Information(3)
                y = rng.Information(6)
                ys.append((ri, pg, y))
            except Exception as e:
                ys.append((ri, None, None))
        diffs = []
        for i, (ri, pg, y) in enumerate(ys):
            if i + 1 < len(ys):
                nri, npg, ny = ys[i+1]
                diff = round(ny - y, 2) if (y is not None and ny is not None and pg == npg) else None
            else:
                diff = None
            print(f"  row{ri:3d} p={pg} y={y} h={diff}")
            if diff is not None and 0 < diff < 60:
                diffs.append(diff)
        if diffs:
            print(f"  median_h={sorted(diffs)[len(diffs)//2]:.2f} n={len(diffs)} min={min(diffs):.2f} max={max(diffs):.2f}")
            # linePitch for b35 = 17.5pt
            snapped_175 = sum(1 for d in diffs if abs(d - 17.5) < 0.5)
            snapped_1475 = sum(1 for d in diffs if abs(d - 14.75) < 0.5)
            print(f"  at 17.5pt (linePitch): {snapped_175}/{len(diffs)}")
            print(f"  at ~14.75pt (unsnapped): {snapped_1475}/{len(diffs)}")
finally:
    wdoc.Close(False)
    word.Quit()
