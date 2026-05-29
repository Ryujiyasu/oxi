"""S445: Measure 7ead52 table row heights in Word vs trHeight values.

vMerge present -> Rows(i) fails. Iterate cells, group by RowIndex,
get each cell's first-para Y (collapsed start, R30) and HeightRule/Height
per-cell (Cell.Row*).
"""
import sys, json, win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
DOCX = r"c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\7ead52b63f0e_000067058.docx"

word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(DOCX, ReadOnly=True)
try:
    tbl = doc.Tables(1)
    rng = tbl.Range
    cells = rng.Cells
    print(f"total cells={cells.Count}")
    rows = {}  # rowindex -> list of (col, y, text, h, rule)
    for ci in range(1, cells.Count + 1):
        cell = cells(ci)
        try:
            ri = cell.RowIndex
            col = cell.ColumnIndex
        except Exception:
            continue
        crng = cell.Range
        startrng = doc.Range(crng.Start, crng.Start)
        try:
            y = round(float(startrng.Information(6)), 3)
        except Exception as e:
            y = None
        try:
            h = round(float(cell.Height), 3)
            rule = int(cell.HeightRule)
        except Exception:
            h, rule = None, None
        rows.setdefault(ri, []).append({
            "col": col, "y": y, "h": h, "rule": rule,
            "text": crng.Text[:14],
        })

    # per-row min Y (topmost text) and reported height
    print("\n--- per row ---")
    rowtops = []
    for ri in sorted(rows):
        cs = rows[ri]
        ys = [c["y"] for c in cs if c["y"] is not None]
        topy = min(ys) if ys else None
        hs = sorted(set(c["h"] for c in cs if c["h"] is not None))
        rules = sorted(set(c["rule"] for c in cs if c["rule"] is not None))
        rowtops.append((ri, topy, hs, rules))
        print(f"row{ri}: topY={topy} heights={hs} rules={rules} ncells={len(cs)}")

    print("\n--- pitch between consecutive row topY ---")
    prev = None
    for ri, topy, hs, rules in rowtops:
        if prev is not None and topy is not None and prev[1] is not None:
            print(f"  row{prev[0]}->row{ri}: {prev[1]:.2f} -> {topy:.2f}  pitch={topy-prev[1]:.2f}  (reported_h={prev[2]})")
        prev = (ri, topy, hs, rules)
finally:
    doc.Close(False)
    word.Quit()
