# -*- coding: utf-8 -*-
"""Whole-document arms of ja/reference__0ea3ec86: change ONE property of the
p1 title «障害者(児)» and read the glyph advances from Word's PDF.

Title: ＭＳ ゴシック 36pt italic, run spacing -16, jc distribute, ind leftChars
800 / rightChars 1800, grid linesAndChars linePitch 411 charSpace 3194.
Word (PDF) advances the kanji by 35.3 (= 36 - 0.8: no grid share) and sets the
title on one line; Oxi gives 42.1 per kanji and wraps «)».

A sliced body (first paragraphs only) made Word wrap every character and COM
Information(5) reads 51.0 for the first glyph even in the original, so the
whole package is kept and the PDF is the instrument.

Arms: orig / nospacing (run spacing removed) / left (jc) / noitalic / sz21.
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SRC = 'pipeline_data/docx_corpus/ja/reference/0ea3ec86480140c2.docx'
OUT = 'tests/fixtures/gridbigchar'; os.makedirs(OUT, exist_ok=True)

z = zipfile.ZipFile(SRC)
doc = z.read('word/document.xml').decode('utf-8')
i = doc.index('(児)</w:t>')
p0 = max(doc.rfind('<w:p>', 0, i), doc.rfind('<w:p ', 0, i))
p1 = doc.index('</w:p>', i) + len('</w:p>')
title = doc[p0:p1]


def write(name, new_title):
    at = os.path.join(OUT, f'whole_{name}.docx')
    xml = doc[:p0] + new_title + doc[p1:]
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as o:
        for it in z.infolist():
            o.writestr(it.filename, xml.encode('utf-8') if it.filename == 'word/document.xml' else z.read(it.filename))
    return at


arms = {
    'orig': title,
    'nospacing': title.replace('<w:spacing w:val="-16"/>', ''),
    'left': title.replace('<w:jc w:val="distribute"/>', '<w:jc w:val="left"/>'),
    'noitalic': title.replace('<w:i/>', ''),
    'sz21': title.replace('w:val="72"', 'w:val="21"'),
    'sp_plus16': title.replace('<w:spacing w:val="-16"/>', '<w:spacing w:val="16"/>'),
    'sp_zero': title.replace('<w:spacing w:val="-16"/>', '<w:spacing w:val="0"/>'),
    'nosp_left': title.replace('<w:spacing w:val="-16"/>', '').replace('<w:jc w:val="distribute"/>', '<w:jc w:val="left"/>'),
    'sp_m2_left': title.replace('<w:spacing w:val="-16"/>', '<w:spacing w:val="-2"/>').replace('<w:jc w:val="distribute"/>', '<w:jc w:val="left"/>'),
}
if len(sys.argv) > 1:
    arms = {k: v for k, v in arms.items() if k in sys.argv[1:]}
files = {k: write(k, v) for k, v in arms.items()}

import win32com.client, fitz
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
try:
    for name, at in files.items():
        pdf = os.path.abspath(at[:-5] + '.pdf')
        d = app.Documents.Open(os.path.abspath(at), ReadOnly=True)
        try:
            d.ExportAsFixedFormat(pdf, 17)
        finally:
            d.Close(False)
        page = fitz.open(pdf)[0]
        rows = []
        for b in page.get_text('rawdict')['blocks']:
            for l in b.get('lines', []):
                chars = [(c['c'], round(c['bbox'][0], 1)) for s in l['spans'] for c in s['chars']]
                if any(c in '障害者児' for c, _ in chars) and l['bbox'][1] < 250:
                    rows.append((round(l['bbox'][1], 1), chars))
        print(f'{name:10} {rows}', flush=True)
finally:
    app.Quit()
