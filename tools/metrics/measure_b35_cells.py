"""Measure each non-merged cell's Y position and height in b35 tables.
Strategy: iterate (row_idx, col_idx) pairs; try tbl.Cell(r,c) and catch
errors for merged cells. Sort by Y to identify unique row Y positions.
"""
import win32com.client, os, sys
sys.stdout.reconfigure(encoding='utf-8')

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx"
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    for ti in range(1, wdoc.Tables.Count + 1):
        tbl = wdoc.Tables(ti)
        n_rows = tbl.Rows.Count
        n_cols = tbl.Columns.Count
        print(f'\n=== Table {ti}: {n_rows}R x {n_cols}C ===')
        # Collect all cell positions via Cell API
        cells = []
        for r in range(1, n_rows + 1):
            for c in range(1, n_cols + 1):
                try:
                    cell = tbl.Cell(r, c)
                    rng = cell.Range
                    pg = rng.Information(3)
                    y = rng.Information(6)
                    x = rng.Information(5)
                    txt = rng.Text.strip()[:20]
                    cells.append((r, c, pg, round(y,2), round(x,2), txt))
                except Exception:
                    # merged continuation
                    pass
        # Get unique Ys per row (row_idx grouping)
        by_row = {}
        for r, c, pg, y, x, txt in cells:
            if r not in by_row:
                by_row[r] = y
        # Compute row-height diffs
        sorted_rows = sorted(by_row.items())
        print(f'  {len(sorted_rows)} unique row-Ys')
        for i in range(len(sorted_rows) - 1):
            r1, y1 = sorted_rows[i]
            r2, y2 = sorted_rows[i+1]
            diff = round(y2 - y1, 2)
            print(f'  row{r1}->{r2}: y {y1:.2f}->{y2:.2f} diff={diff}')
finally:
    wdoc.Close(False)
    word.Quit()
