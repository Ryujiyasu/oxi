# -*- coding: utf-8 -*-
"""RENDER-TRUTH: export ikujidetail to PDF via Word, fitz line bboxes on p1.
Determines the TRUE rendered line heights (gold standard) to resolve the
COM-table (16.5/21.0) vs Information(6)-gap (14.2/18.0) contradiction."""
import os, sys, time
sys.stdout.reconfigure(encoding='utf-8')
DOCX = os.path.abspath('tools/golden-test/documents/docx/ikujidetail_002197815.docx')
PDF  = r'C:\Users\ryuji\AppData\Local\Temp\ikd_truth.pdf'
if not os.path.exists(PDF) or '--reexport' in sys.argv:
    import win32com.client as win32
    w = win32.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(DOCX, ReadOnly=True)
        d.ExportAsFixedFormat(PDF, 17)  # 17 = wdExportFormatPDF
        d.Close(False)
    finally:
        w.Quit()
    print('exported PDF')
import fitz
doc = fitz.open(PDF)
pg = doc[0]
lines = []
for blk in pg.get_text('dict')['blocks']:
    if blk.get('type', 0) != 0: continue
    for ln in blk.get('lines', []):
        txt = ''.join(s['text'] for s in ln['spans']).strip()
        if not txt: continue
        y0 = min(s['bbox'][1] for s in ln['spans'])
        sz = max(s['size'] for s in ln['spans'])
        lines.append((round(y0,2), round(sz,1), txt[:26]))
lines.sort()
print('  %-8s %-5s %-6s %s'%('y0','size','pitch','text'))
prev=None
for y,sz,t in lines[:16]:
    p = round(y-prev,2) if prev is not None else 0
    print('  %-8.2f %-5.1f %-6.2f %s'%(y,sz,p,t))
    prev=y
