"""Measure 0e7a p2 paragraph Y positions in Word."""
import sys, os, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx"
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    print(f"Paragraph count: {wdoc.Paragraphs.Count}")
    for i in range(1, wdoc.Paragraphs.Count + 1):
        p = wdoc.Paragraphs(i)
        rng = p.Range
        pg = rng.Information(3)
        if pg != 2: continue
        y = rng.Information(6)
        txt = rng.Text[:40].replace('\r', '\\r').replace('\n', '\\n')
        print(f"  para{i:3d} p{pg} y={y:7.2f} [{txt}]")
finally:
    wdoc.Close(False)
    word.Quit()
