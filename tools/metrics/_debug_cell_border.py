"""Debug: probe one fixture with multiple measurement methods."""
import time
import win32com.client

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
time.sleep(2.0)

path = r"C:\Users\ryuji\oxi-2\tools\metrics\output\cell_border_absorption_fixtures\CBA_sz12hp_(6.0pt).docx"
wdoc = word.Documents.Open(path)
try:
    wdoc.Repaginate()
    time.sleep(0.1)

    tbl = wdoc.Tables(1)
    print(f"n_rows={tbl.Rows.Count}, n_cols={tbl.Columns.Count}")

    cell = tbl.Cell(1, 1)
    print(f"\nCell(1,1):")
    print(f"  Range.Text={cell.Range.Text!r}")

    # Try LeftPadding with various access patterns
    try:
        lp = cell.LeftPadding
        print(f"  LeftPadding={lp}")
    except Exception as e:
        print(f"  LeftPadding ERR: {e}")

    try:
        b = cell.Borders(1)
        print(f"  Borders(1).LineStyle={b.LineStyle}, LineWidth={b.LineWidth}")
    except Exception as e:
        print(f"  Borders ERR: {e}")

    # Range positions
    print(f"  Range.Information(5)={cell.Range.Information(5)} (cell content X)")
    print(f"  Range.Information(6)={cell.Range.Information(6)} (cell content Y)")

    # First text character
    text = cell.Range.Text
    idx = text.find("R")
    print(f"  text.find('R')={idx}")
    if idx >= 0:
        sub = wdoc.Range(cell.Range.Start + idx, cell.Range.Start + idx + 1)
        print(f"  Range[R].Information(5)={sub.Information(5)}")

    # Selection-based
    sel = word.Selection
    if idx >= 0:
        sel.SetRange(cell.Range.Start + idx, cell.Range.Start + idx + 1)
        print(f"  Selection.Information(5)={sel.Information(5)}")
        print(f"  Selection.Information(8)={sel.Information(8)} (wdHorizontalPositionRelativeToTextBoundary)")

    # Word's table-level cell margin via COM (TableNormal.LeftPadding maybe)
    try:
        lc = tbl.LeftPadding
        print(f"  Table.LeftPadding={lc}")
    except Exception as e:
        print(f"  Table.LeftPadding ERR: {e}")

    # Try cell.WordWrap
    print(f"  Cell.Width={cell.Width}")

finally:
    wdoc.Close(False)
    word.Quit()
