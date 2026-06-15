# -*- coding: utf-8 -*-
"""tokyoshugyo render-truth: per-page Word TEXT-LINE count (PDF via
ExportAsFixedFormat + fitz) AND horizontal table-border Y positions, compared
to Oxi's GDI --dump-layout. Localizes which pages Oxi over-packs (the source of
the doc-wide -1 cell-height drift). Word page N's line count vs Oxi page N's.

Usage: python tools/metrics/_tks_pdf_lines.py [--reexport] [--redump]
"""
import os, sys, json, subprocess, tempfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', 'tokyoshugyo_000599795.docx')
PDF  = os.path.join(tempfile.gettempdir(), 'tks_truth.pdf')
DUMP = os.path.join(tempfile.gettempdir(), 'tks_oxi_dump.json')
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

if not os.path.exists(PDF) or '--reexport' in sys.argv:
    import win32com.client as win32
    w = win32.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(DOCX, ReadOnly=True)
        d.ExportAsFixedFormat(PDF, 17)  # wdExportFormatPDF
        d.Close(False)
    finally:
        w.Quit()
    print('exported PDF')

if not os.path.exists(DUMP) or '--redump' in sys.argv:
    subprocess.run([RENDERER, DOCX, os.path.join(tempfile.gettempdir(), 'tks_p_'),
                    '--dump-layout=' + DUMP], capture_output=True, timeout=180)
    print('dumped Oxi layout')

import fitz
doc = fitz.open(PDF)

def page_text_lines(pg):
    ys = []
    for blk in pg.get_text('dict')['blocks']:
        if blk.get('type', 0) != 0: continue
        for ln in blk.get('lines', []):
            txt = ''.join(s['text'] for s in ln['spans']).strip()
            if not txt: continue
            y0 = min(s['bbox'][1] for s in ln['spans'])
            if y0 < 60 or y0 > 800: continue  # header/footer
            ys.append(round(y0, 1))
    return sorted(ys)

word_lc = {}
for pi in range(len(doc)):
    word_lc[pi+1] = len(page_text_lines(doc[pi]))

# Oxi: per page, count distinct text-line Y rows
d = json.load(open(DUMP, encoding='utf-8'))
oxi_lc = {}
for i, pg in enumerate(d['pages']):
    ys = set()
    for el in pg.get('elements', []):
        if el.get('type') == 'text':
            ys.add(round(el['y'], 0))
    oxi_lc[i+1] = len(ys)

print(f"Word pages: {len(doc)}  Oxi pages: {len(d['pages'])}")
print(f"{'pg':>3} {'Word':>5} {'Oxi':>5} {'diff':>5}  (Oxi-Word; +=Oxi packs more)")
cum = 0
for p in range(1, max(len(doc), len(d['pages']))+1):
    wl = word_lc.get(p, 0); ol = oxi_lc.get(p, 0)
    diff = ol - wl
    cum += diff
    flag = ''
    if diff > 0: flag = ' <-- Oxi over'
    elif diff < 0: flag = ' (Oxi under)'
    print(f"{p:>3} {wl:>5} {ol:>5} {diff:>+5}  cum={cum:>+4}{flag}")
