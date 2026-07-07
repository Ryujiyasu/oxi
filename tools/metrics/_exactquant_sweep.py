# -*- coding: utf-8 -*-
"""Controlled sweep: Word's EXACT-lineRule cursor quantization (the bd90b00
備考 knife-edge's largest component). Runs of identical exact-N empty
paragraphs bracketed by text anchors; COM Information(6) per paragraph.

FINDING (2026-07-07): steady advances oscillate N / N+0.05 with ONE −0.6
outlier per run (phase-dependent; the 4th of five in bd90b00, the 1st here);
a 10-line run sums to nominal −0.3 — Word accumulates exact-line cursor
positions on a device quantum (the S629 render-snap phenomenon at the
CURSOR level). Info6's 0.05 rounding hides the exact law: derive it from
PDF-fine anchor positions (the text paras bracketing the runs) × run length
× exact values.

Run: python tools/metrics/_exactquant_sweep.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

esc = pg.esc
MINCHO = pg.MINCHO
OUTDIR = os.path.join(os.environ.get("TEMP", "."), 'gridquant')
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, 'exactquant.docx')

def rpr(sz="21"):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{sz}"/>')

def para(txt, spacing=''):
    r = rpr()
    return (f'<w:p><w:pPr>{spacing}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

def emptyp(spacing):
    r = rpr()
    return f'<w:p><w:pPr>{spacing}<w:rPr>{r}</w:rPr></w:pPr></w:p>'

EX300 = '<w:spacing w:line="300" w:lineRule="exact"/>'
EX200 = '<w:spacing w:line="200" w:lineRule="exact"/>'

body = [para('先頭', EX200)]
for _ in range(10):
    body.append(emptyp(EX300))
body.append(para('中間', EX200))
for _ in range(10):
    body.append(emptyp(EX200))
body.append(para('末尾', EX200))
SECTPR = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="1134" w:right="1134" w:bottom="851" w:left="1134" w:header="851" w:footer="567" w:gutter="0"/>'
          '<w:docGrid w:type="lines" w:linePitch="330"/></w:sectPr>')
pg.write_docx(DOCX, pg.doc(''.join(body) + SECTPR))
print('wrote', DOCX)

import win32com.client
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
sys.stdout.reconfigure(encoding='utf-8')
prev = None
for i in range(1, doc.Paragraphs.Count + 1):
    p = doc.Paragraphs(i)
    rng = doc.Range(p.Range.Start, p.Range.Start)
    y = rng.Information(6)
    adv = '' if prev is None else f' adv={y-prev:+.3f}'
    prev = y
    print(f'i={i:2d} y={y:8.2f}{adv}  {repr(p.Range.Text.strip()[:6])}')
doc.Close(False)
word.Quit()
