"""Measure Word table positions in 683f to determine border overhead.

Goal: Find why Oxi tables are 2px (1.5pt) shorter than Word's.
The 683f doc has 2 tables (1 row, 1 cell each) with all borders.
Oxi assumes outer borders add no row-height overhead.
"""
import win32com.client, time, sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/683ffcab86e2_20230331_resources_open_data_contract_addon_00.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = True
word.DisplayAlerts = False
doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(1)

print(f"Total tables: {doc.Tables.Count}")

# wdVerticalPositionRelativeToPage = 6
WD_Y_PAGE = 6

for ti in range(1, doc.Tables.Count + 1):
    t = doc.Tables(ti)
    print(f"\n=== Table {ti} ===")
    print(f"  Rows: {t.Rows.Count}")
    # Top of table
    top_y = t.Range.Information(WD_Y_PAGE)
    # Bottom: get last cell range end
    last_cell = t.Cell(t.Rows.Count, t.Columns.Count)
    last_y = last_cell.Range.Information(WD_Y_PAGE)
    print(f"  Top Y: {top_y:.2f}pt")
    print(f"  Last cell Y: {last_y:.2f}pt")
    # Row heights
    for ri in range(1, t.Rows.Count + 1):
        r = t.Rows(ri)
        try:
            h = r.Height
            hr = r.HeightRule
            print(f"  Row {ri}: Height={h}, HeightRule={hr}")
        except Exception as e:
            print(f"  Row {ri}: err={e}")
        # Cell content top/bottom Y
        cell = t.Cell(ri, 1)
        cell_top_y = cell.Range.Information(WD_Y_PAGE)
        # Get last char of cell
        try:
            chars = cell.Range.Characters
            n = chars.Count
            if n > 0:
                last_char_y = chars(n).Information(WD_Y_PAGE)
                print(f"    cell content Y range: {cell_top_y:.2f} - {last_char_y:.2f}")
        except Exception as e:
            print(f"    err: {e}")
    # Borders
    print(f"  Borders:")
    for bi, name in [(-1, "Top"), (-2, "Left"), (-3, "Bottom"), (-4, "Right"),
                      (-5, "InsideH"), (-6, "InsideV")]:
        try:
            b = t.Borders(bi)
            print(f"    {name}: LineStyle={b.LineStyle} LineWidth={b.LineWidth}")
        except Exception as e:
            print(f"    {name}: err")

# Measure paragraph Y between tables and after
print("\n=== All paragraphs Y (first 50) ===")
for pi in range(1, min(60, doc.Paragraphs.Count + 1)):
    p = doc.Paragraphs(pi)
    y = p.Range.Information(WD_Y_PAGE)
    in_cell = p.Range.Information(12)  # wdWithInTable
    txt = p.Range.Text[:30].replace('\r', '\\r').replace('\x07', '\\BEL')
    marker = "  [TBL]" if in_cell else ""
    print(f"  P{pi}: y={y:.2f} {marker} '{txt}'")

doc.Close(SaveChanges=False)
word.Quit()
