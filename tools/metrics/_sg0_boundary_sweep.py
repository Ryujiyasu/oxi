# -*- coding: utf-8 -*-
"""Controlled sweep: the snapped -> snapToGrid=0 boundary conventions in a
typed docGrid (the bd90b00/1636 備考-knife-edge family root).

Measures per config: [3 snapped 10.5pt paras] [X] [table] where X varies:
  - none:      no para between (snapped -> table reference)
  - sg0:       one snapToGrid=0 auto para (the 以下の各事項 pattern)
  - sg0x2:     two sg0 paras (sg0 -> sg0 internal advance)
  - snapped:   one more snapped para (control)
Quantities: last-snapped ink -> X ink (start convention), X ink -> X2/table
border (X's own advance), at pitch 330 and 360.

Run: python tools/metrics/_sg0_boundary_sweep.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), "gridquant")
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, "sg0b.docx")
PDF = os.path.join(OUTDIR, "sg0b.pdf")

esc = pg.esc
MINCHO = pg.MINCHO

def rpr(szhalf="21"):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{szhalf}"/>')

def para(txt, ppr_extra=''):
    r = rpr()
    return (f'<w:p><w:pPr>{ppr_extra}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

SG0 = '<w:snapToGrid w:val="0"/><w:spacing w:line="240" w:lineRule="auto"/>'

def table():
    return ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{para("表内")}</w:tc></w:tr></w:tbl>')

def pagebreak_para():
    r = rpr()
    return (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:br w:type="page"/></w:r></w:p>')

def sect(pitch, last=False):
    inner = (f'<w:pgSz w:w="11906" w:h="16838"/>'
             f'<w:pgMar w:top="1134" w:right="1134" w:bottom="851" w:left="1134" w:header="851" w:footer="567" w:gutter="0"/>'
             f'<w:docGrid w:type="lines" w:linePitch="{pitch}"/>')
    if last:
        return f'<w:sectPr>{inner}</w:sectPr>'
    return f'<w:p><w:pPr><w:sectPr>{inner}</w:sectPr></w:pPr></w:p>'

S1, S2, S3 = 'あ一', 'あ二', 'あ三'
X1, X2 = 'えっくす壱', 'えっくす弐'

VARIANTS = ['none', 'sg0', 'sg0x2', 'snapped']
PITCHES = [330, 360]
configs = []
body = []
for pi_, pitch in enumerate(PITCHES):
    for v in VARIANTS:
        if body:
            body.append(pagebreak_para())
        configs.append((pitch, v))
        body.append(para(S1))
        body.append(para(S2))
        body.append(para(S3))
        if v == 'sg0':
            body.append(para(X1, SG0))
        elif v == 'sg0x2':
            body.append(para(X1, SG0))
            body.append(para(X2, SG0))
        elif v == 'snapped':
            body.append(para(X1))
        body.append(table())
    if pi_ + 1 < len(PITCHES):
        body.append(sect(pitch))
body.append(sect(PITCHES[-1], last=True))

pg.write_docx(DOCX, pg.doc(''.join(body)))
print('wrote', DOCX)

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
# realign by content: a config page contains あ一
res = []
for p in range(len(d)):
    page = d[p]
    marks = {}
    for b in page.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            txt = ''.join(s['text'] for s in l.get('spans', []))
            for key in (S1, S2, S3, X1, X2, '表内'):
                if key in txt.replace(' ', ''):
                    marks[key] = round(l['bbox'][1], 2)
    if S1 in marks:
        ys = set()
        for dr in page.get_drawings():
            for it in dr['items']:
                if it[0] == 'l':
                    p1, p2 = it[1], it[2]
                    if abs(p1.y-p2.y) < 0.2 and abs(p1.x-p2.x) > 60:
                        ys.add(round(p1.y, 2))
                elif it[0] == 're':
                    rr = it[1]
                    if rr.height < 2 and rr.width > 60:
                        ys.add(round(rr.y0, 2))
        merged = []
        for y in sorted(ys):
            if merged and abs(merged[-1]-y) < 1.0:
                continue
            merged.append(y)
        res.append((marks, merged))
for (pitch, v), (marks, borders) in zip(configs, res):
    s3 = marks.get(S3)
    x1 = marks.get(X1)
    x2 = marks.get(X2)
    tb = borders[0] if borders else None
    parts = [f'pitch={pitch/20:5.2f} {v:8s} S3={s3}']
    if x1 is not None:
        parts.append(f'S3->X1={round(x1-s3,2)}')
        if x2 is not None:
            parts.append(f'X1->X2={round(x2-x1,2)}')
            parts.append(f'X2->tbl={round(tb-x2,2) if tb else None}')
        else:
            parts.append(f'X1->tbl={round(tb-x1,2) if tb else None}')
    else:
        parts.append(f'S3->tbl={round(tb-s3,2) if tb else None}')
    print('  '.join(parts))
