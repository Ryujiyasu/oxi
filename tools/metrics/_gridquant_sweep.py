# -*- coding: utf-8 -*-
"""Controlled sweep: Word's typed-docGrid line quantization as a function of
(linePitch, fontSize). Resolves the 2ea81a grid323 2-cell question and the
deferred probeqsizes finding (28pt = 39.6/line, 6pt = 11/line at pitch 360 —
neither is a cell multiple).

One SECTION per pitch (sectPr chain), each section = one page per fs config:
marker + 5 identical single-line MS-Mincho paras (no spacing overrides,
snapped). Steady-state para-to-para ink gap (paras 3-4, 4-5) = the quantized
advance for that (pitch, fs). A leading odd-height para (exact-170 = 8.5pt)
variant per fs measures the PHASE behavior (absolute-slot check).

Run: python tools/metrics/_gridquant_sweep.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), "gridquant")
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, "gridquant.docx")
PDF = os.path.join(OUTDIR, "gridquant.pdf")

esc = pg.esc
MINCHO = pg.MINCHO

PITCHES = [312, 323, 330, 360, 480]
FSHALF = [12, 16, 21, 24, 28, 32, 42, 56]  # 6,8,10.5,12,14,16,21,28 pt

def rpr(szhalf):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{szhalf}"/>')

def para(txt, szhalf, spacing=''):
    r = rpr(szhalf)
    return (f'<w:p><w:pPr>{spacing}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')

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

MARK = 'あいう'
configs = []  # (pitch, fs_half, phase)
body = []
for pi_, pitch in enumerate(PITCHES):
    for j, fs in enumerate(FSHALF):
        for phase in (0, 1):
            if body:
                body.append(pagebreak_para())
            configs.append((pitch, fs, phase))
            if phase:
                # odd-height leader (exact 170tw = 8.5pt) to knock the cursor off-phase
                body.append(para('P', "16", '<w:spacing w:line="170" w:lineRule="exact"/>'))
            for k in range(5):
                body.append(para(MARK * 3, str(fs)))
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
print('pages:', len(d))
sys.stdout.reconfigure(encoding='utf-8')
print(f'{"pitch":>5} {"fs":>5} {"ph":>2} {"gaps (para ink-to-ink)":<40} natural')
for i, (pitch, fs, phase) in enumerate(configs):
    if i >= len(d):
        break
    page = d[i]
    ys = []
    for b in page.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            txt = ''.join(s['text'] for s in l.get('spans', []))
            if MARK[0] in txt:
                ys.append(round(l['bbox'][1], 2))
    ys.sort()
    gaps = [round(ys[k+1]-ys[k], 2) for k in range(len(ys)-1)]
    fspt = fs / 2.0
    nat = round(fspt * 83.0 / 64.0, 2)
    print(f'{pitch/20:5.2f} {fspt:5.1f} {phase:2d} {str(gaps):<40} nat={nat} first={ys[0] if ys else None}')
