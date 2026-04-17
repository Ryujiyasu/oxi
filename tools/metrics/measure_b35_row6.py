"""Measure b35 table 1 row 6 in detail: why Oxi renders 35.5pt but Word 20.45pt.
Row 6 is 人的管理措置 category intro row."""
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
    print(f"Table 1: {tbl.Rows.Count} rows x {tbl.Columns.Count} cols")

    # All rows 5-8 detail
    for r in range(5, 9):
        print(f"\n=== Row {r} ===")
        for c in range(1, 3):
            try:
                cell = tbl.Cell(r, c)
            except Exception as e:
                print(f"  col{c}: err {e}")
                continue
            rng = cell.Range
            y = rng.Information(6)
            txt = rng.Text.replace('\r', '\\r').replace('\n', '\\n')[:80]
            n_paras = cell.Range.Paragraphs.Count
            try: w = cell.Width
            except: w = '?'
            try: h = cell.Height
            except: h = '?'
            print(f"  col{c}: y={y:.2f} w={w} h={h} paras={n_paras}")
            print(f"    text={txt!r}")
            for i in range(1, n_paras + 1):
                try:
                    p = cell.Range.Paragraphs(i)
                    pt = p.Range.Text[:50].replace('\r', '\\r')
                    fmt = p.Format
                    print(f"    para{i}: ls={fmt.LineSpacing:.1f} rule={fmt.LineSpacingRule} sa={fmt.SpaceAfter:.1f} sb={fmt.SpaceBefore:.1f} text={pt!r}")
                except Exception as e:
                    print(f"    para{i}: err {e}")
finally:
    wdoc.Close(False)
    word.Quit()
