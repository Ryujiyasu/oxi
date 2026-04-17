"""Count Word's paragraphs per page for 0e7a."""
import sys, os, win32com.client
sys.stdout.reconfigure(encoding='utf-8')

DOC = r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx"
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
wdoc = word.Documents.Open(os.path.abspath(DOC), ReadOnly=True)
try:
    wdoc.Repaginate()
    n_paras = wdoc.Paragraphs.Count
    print(f"Total paras: {n_paras}")
    pages = {}
    for i in range(1, n_paras + 1):
        p = wdoc.Paragraphs(i)
        pg = p.Range.Information(3)
        pages.setdefault(pg, []).append(i)
    for pg in sorted(pages.keys()):
        paras = pages[pg]
        print(f'Page {pg}: paras {paras[0]}-{paras[-1]} ({len(paras)} paras)')
    # Print first+last para on each page's content
    print()
    for pg in sorted(pages.keys())[:3]:
        paras = pages[pg]
        print(f'=== Page {pg} ===')
        for i in paras[:3] + paras[-3:]:
            try:
                p = wdoc.Paragraphs(i)
                txt = p.Range.Text[:40].replace('\r', '\\r')
                y = p.Range.Information(6)
                print(f'  para{i}: y={y:7.2f} [{txt}]')
            except: pass
finally:
    wdoc.Close(False)
    word.Quit()
