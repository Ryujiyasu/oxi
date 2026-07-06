# -*- coding: utf-8 -*-
"""Controlled sweep: Word's typed-docGrid EMPTY-paragraph cell count as a
function of (linePitch, pilcrow fontSize). The S583/S195 family at pitches
beyond 360: 2ea81a (pitch 323, sz28 empty) renders 1 CELL in Word while
kojin (pitch 360, sz28 empty, CJK ascii) measured 2 CELLS (S580/S583).

Layout per config page: anchor text para (10.5pt) / 3 EMPTY paras (fs) /
anchor text para (10.5pt). Empty height = (ink2 - ink1 - anchor_line) / 3.

Run: python tools/metrics/_emptyquant_sweep.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), "gridquant")
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, "emptyquant.docx")
PDF = os.path.join(OUTDIR, "emptyquant.pdf")

esc = pg.esc
MINCHO = pg.MINCHO

PITCHES = [323, 330, 360]
FSHALF = [21, 24, 28, 32]  # pilcrow size of the EMPTY paras

def rpr(szhalf):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{szhalf}"/>')

def para(txt, szhalf):
    r = rpr(szhalf)
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

def emptyp(szhalf):
    r = rpr(szhalf)
    return f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr></w:p>'

def pagebreak_para():
    r = rpr("21")
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:br w:type="page"/></w:r></w:p>')

def sect(pitch, last=False):
    inner = (f'<w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1134" w:right="1134" w:bottom="851" w:left="1134" w:header="851" w:footer="567" w:gutter="0"/>'
             f'<w:docGrid w:type="lines" w:linePitch="{pitch}"/>')
    if last:
        return f'<w:sectPr>{inner}</w:sectPr>'
    return f'<w:p><w:pPr><w:sectPr>{inner}</w:sectPr></w:pPr></w:p>'

A1 = 'あ上'
A2 = 'あ下'
configs = []
body = []
for pi_, pitch in enumerate(PITCHES):
    for fs in FSHALF:
        if body:
            body.append(pagebreak_para())
        configs.append((pitch, fs))
        body.append(para(A1, "21"))
        for _ in range(3):
            body.append(emptyp(str(fs)))
        body.append(para(A2, "21"))
    if pi_ + 1 < len(PITCHES):
        body.append(sect(pitch))
body.append(sect(PITCHES[-1], last=True))

pg.write_docx(DOCX, pg.doc(''.join(body)))
print('wrote', DOCX, 'configs:', len(configs))

import win32com.client
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
doc.ExportAsFixedFormat(PDF, 17)
doc.Close(False)
word.Quit()
print('exported', PDF)

import fitz
d = fitz.open(PDF)
sys.stdout.reconfigure(encoding='utf-8')
print('pages:', len(d))
# realign: pages with BOTH anchors are config pages, in order
results = []
for p in range(len(d)):
    page = d[p]
    y1 = y2 = None
    for b in page.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            txt = ''.join(s['text'] for s in l.get('spans', []))
            if A1 in txt:
                y1 = round(l['bbox'][1], 2)
            elif A2 in txt:
                y2 = round(l['bbox'][1], 2)
    if y1 is not None and y2 is not None:
        results.append((p, y1, y2))
print(f'{"pitch":>6} {"fs":>5} {"gap":>7} {"3empties":>9} {"per-empty":>9} {"cells":>6}')
for (pitch, fs), (p, y1, y2) in zip(configs, results):
    ppt = pitch / 20.0
    # anchor line (10.5pt at this pitch) = 1 cell
    total = y2 - y1
    three = total - ppt  # minus the top anchor's own line (1 cell)
    per = three / 3.0
    print(f'{ppt:6.2f} {fs/2:5.1f} {total:7.2f} {three:9.2f} {per:9.2f} {per/ppt:6.2f}')
