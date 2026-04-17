"""Measure b35 table 1 row 1 in detail to understand content vs row height."""
import sys, os, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx"
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    tbl = wdoc.Tables(1)
    n_cols = tbl.Columns.Count
    n_rows = tbl.Rows.Count
    print(f"Table 1: {n_rows}R x {n_cols}C")

    # Column widths (measured from first row if all cells present)
    widths = []
    for c in range(1, n_cols + 1):
        try:
            cell = tbl.Cell(1, c)
            widths.append((c, cell.Width))
        except:
            try:
                cell = tbl.Cell(2, c)
                widths.append((c, cell.Width))
            except: widths.append((c, None))
    print("Column widths (pt):")
    for c, w in widths:
        print(f"  col{c}: {w}")

    # Row 1 cells
    print("\nRow 1 cell details:")
    for c in range(1, n_cols + 1):
        try:
            cell = tbl.Cell(1, c)
        except: continue
        rng = cell.Range
        y = rng.Information(6)
        txt = rng.Text.strip()[:40]
        n_paras = cell.Range.Paragraphs.Count
        w = cell.Width
        print(f"  cell({c}) y={y:.2f} w={w:.1f} paras={n_paras} text={txt!r}")
        for i in range(1, n_paras + 1):
            try:
                p = cell.Range.Paragraphs(i)
                pt = p.Range.Text.strip()[:30]
                # Compute line count by Sentences or actually lines: use Range.Information(10)?
                # Info 10 = line numbered
                try:
                    first_line = p.Range.Characters(1).Information(10)
                    last_line = p.Range.Characters(p.Range.Characters.Count).Information(10)
                    nlines = last_line - first_line + 1
                except Exception as e:
                    nlines = "err"
                # Line spacing
                fmt = p.Format
                print(f"    para{i}: lines={nlines} ls={fmt.LineSpacing} rule={fmt.LineSpacingRule} text={pt!r}")
            except Exception as e:
                print(f"    para{i}: err {e}")

    # Row 2 cells for comparison
    print("\nRow 2 cell details:")
    for c in range(1, min(3, n_cols + 1)):
        try:
            cell = tbl.Cell(2, c)
        except: continue
        rng = cell.Range
        y = rng.Information(6)
        txt = rng.Text.strip()[:40]
        n_paras = cell.Range.Paragraphs.Count
        print(f"  cell({c}) y={y:.2f} paras={n_paras} text={txt!r}")

finally:
    wdoc.Close(False)
    word.Quit()
