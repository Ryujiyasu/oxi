"""COM-measure b35 table 1 row 1 cell "区分" line height and compare to
Oxi's word_line_height_table_cell / word_line_height_standard for MS Mincho 10.5pt.

Goal: determine which line-height function estimate_para_height SHOULD use
when adjust_line_height_in_table=true.
"""
import sys, os, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx"
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    # adjust_line_height_in_table compat flag (12)
    try: al = wdoc.Compatibility(12)
    except: al = "err"
    print(f"adjust_line_height_in_table (compat 12): {al}")

    tbl = wdoc.Tables(1)
    print(f"Table 1: {tbl.Rows.Count} rows")

    # Row 1 col 1 "区分"
    cell = tbl.Cell(1, 1)
    para = cell.Range.Paragraphs(1)
    fmt = para.Format
    print(f"  Cell(1,1) text='{cell.Range.Text.strip()[:20]}'")
    print(f"    LineSpacing={fmt.LineSpacing:.3f}pt  Rule={fmt.LineSpacingRule}")
    print(f"    SpaceBefore={fmt.SpaceBefore:.3f}  SpaceAfter={fmt.SpaceAfter:.3f}")

    # Measure cell's rendered height: top Y of cell 1,1 vs cell 2,1
    y1 = tbl.Cell(1, 1).Range.Information(6)
    y2 = tbl.Cell(2, 1).Range.Information(6)
    print(f"  Cell(1,1) y={y1:.3f}  Cell(2,1) y={y2:.3f}")
    print(f"  ROW 1 HEIGHT = {y2 - y1:.3f}pt")

    # Cell padding — default is 55tw=2.75pt left/right, 0 top/bottom
    print(f"  Cell(1,1).TopPadding={cell.TopPadding:.3f} BottomPadding={cell.BottomPadding:.3f}")
    print(f"  Cell(1,1).LeftPadding={cell.LeftPadding:.3f} RightPadding={cell.RightPadding:.3f}")

    # Table-level default margins
    tbl_fmt = tbl
    try:
        print(f"  Table.TopPadding={tbl.TopPadding:.3f} BottomPadding={tbl.BottomPadding:.3f}")
    except: pass

    # Run the cell's first paragraph text and measure its font
    rng = cell.Range
    # Get first char's font
    try:
        first_char = rng.Characters(1)
        font_name = first_char.Font.NameFarEast or first_char.Font.Name
        font_size = first_char.Font.Size
        print(f"  Text font: {font_name} {font_size}pt")
    except Exception as e:
        print(f"  Font error: {e}")
finally:
    wdoc.Close(False)
    word.Quit()
