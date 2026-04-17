"""Measure Word's p2 last paragraphs to compare with Oxi p2.

Output: what's the last para on Word's p2? What's its COM para index and text?
Compare to Oxi p2 last para_idx=25 ("・..." at y=686-740).
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
    # Collect all body paragraphs on p2
    p2_paras = []
    for pi, p in enumerate(doc.Paragraphs, 1):
        try:
            pg = p.Range.Information(3)
            y = p.Range.Information(6)
            if pg == 2:
                txt = p.Range.Text.replace('\r','').replace('\x07','')[:50]
                p2_paras.append((pi, y, txt))
        except Exception:
            pass
    p2_paras.sort(key=lambda r: r[1])
    print(f'Word p2: {len(p2_paras)} paragraphs')
    # Show last 10
    for (pi, y, txt) in p2_paras[-10:]:
        print(f'  para_idx={pi:3d} y={y:6.1f} text={txt!r}')

    # Also show first paragraph on p3 (to pinpoint the boundary)
    print('\nWord p3 first 3 paragraphs:')
    p3_paras = []
    for pi, p in enumerate(doc.Paragraphs, 1):
        try:
            pg = p.Range.Information(3)
            y = p.Range.Information(6)
            if pg == 3:
                txt = p.Range.Text.replace('\r','').replace('\x07','')[:50]
                p3_paras.append((pi, y, txt))
        except Exception:
            pass
    p3_paras.sort(key=lambda r: r[1])
    for (pi, y, txt) in p3_paras[:5]:
        print(f'  para_idx={pi:3d} y={y:6.1f} text={txt!r}')

    doc.Close(False)
finally:
    word.Quit()
