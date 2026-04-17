"""Measure Y position of each paragraph and page for b35."""
import sys, os, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx"
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    print(f"Paragraph count: {wdoc.Paragraphs.Count}")
    print(f"Page count (Info 4 on last char): {wdoc.Range().Characters(wdoc.Range().Characters.Count).Information(4)}")
    for i in range(1, min(wdoc.Paragraphs.Count, 40) + 1):
        p = wdoc.Paragraphs(i)
        rng = p.Range
        pg = rng.Information(3)
        y_pt = rng.Information(6)
        x_pt = rng.Information(5)
        txt = rng.Text[:40].replace('\r', '\\r').replace('\n', '\\n')
        print(f"  para{i:2d} p{pg} y={y_pt:7.2f} x={x_pt:7.2f}  [{txt}]")
finally:
    wdoc.Close(False)
    word.Quit()
