"""Measure Word's p3 content BEFORE Table 1 to find drift origin.

Hypothesis: Oxi's Table 1 starts at y=233.5, Word's at ~216.5. The 17pt
drift comes from accumulated differences in content above the table on p3.
"""
import os, sys, time
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
try:
    doc = word.Documents.Open(DOC, ReadOnly=True); time.sleep(0.3)
    doc.Repaginate()
    # Find all paragraphs on page 3 BEFORE the first table
    tbl_start_y = doc.Tables(1).Cell(1,1).Range.Information(6)
    print(f'Word Table 1 cell starts at y={tbl_start_y}')
    # Walk body paragraphs; collect those on p3
    p3_paras = []
    for pi, p in enumerate(doc.Paragraphs, 1):
        try:
            pg = p.Range.Information(3)
            y = p.Range.Information(6)
            if pg == 3:
                txt = p.Range.Text.replace('\r','').replace('\x07','')[:40]
                p3_paras.append((pi, y, txt))
        except Exception:
            pass
    # Sort by y, filter to before-table
    p3_paras.sort(key=lambda r: r[1])
    print(f'\nWord p3 paragraphs ({len(p3_paras)} total, those before Table 1):')
    for (pi, y, txt) in p3_paras:
        marker = ' <<< TABLE' if abs(y - tbl_start_y) < 1 else ''
        if y <= tbl_start_y + 0.5:
            print(f'  para_idx={pi} y={y:.1f}{marker} text={txt!r}')
    doc.Close(False)
finally:
    word.Quit()
