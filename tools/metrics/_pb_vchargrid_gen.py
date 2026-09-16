# -*- coding: utf-8 -*-
"""Vertical (tbRl) text under a `docGrid type="snapToChars"`: what character advance and
column pitch does Word use?

creative__3d2bf04c (landscape, tbRl, docGrid snapToChars linePitch=671 charSpace=49418, ＭＳ 明朝
10.5): Word fits far fewer characters per column and columns per page than Oxi (pcd −3);
Oxi honours the vertical char grid only behind OXI_VERTICAL_CHAR_GRID.

Sheet: vertical section with the grid, one paragraph of 40 kana + a second paragraph. Arms:
charSpace × linePitch × font size. Readout: per-character y advance (Information 6) for the
first 8 characters, column x (Information 5) of paragraphs 1 and 2, characters per column.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/vchargrid'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
KANA = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん'


def document(sz, line_pitch, char_space):
    rpr = f'<w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:hint="eastAsia"/><w:sz w:val="{sz}"/></w:rPr>'
    body = (f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>{KANA[:40]}</w:t></w:r></w:p>'
            f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>{KANA[:10]}</w:t></w:r></w:p>')
    sect = ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1700" w:bottom="1440" w:left="1700" w:header="0" w:footer="0" w:gutter="0"/>'
            f'<w:cols w:space="425"/><w:textDirection w:val="tbRl"/><w:docGrid w:type="snapToChars" w:linePitch="{line_pitch}" w:charSpace="{char_space}"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = [(21, 671, 49418), (21, 360, 49418), (21, 671, 8192), (21, 671, 0), (24, 671, 49418), (20, 671, 49418), (21, 500, 20480)]
    for sz, lp, cs in arms:
        at = OUT / f'sz{sz}_lp{lp}_cs{cs}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/document.xml', document(sz, lp, cs))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            p1 = d.Paragraphs(1).Range; p2 = d.Paragraphs(2).Range
            ys = [round(d.Range(c, c).Information(6), 2) for c in range(p1.Start, p1.Start + 9)]
            adv = [round(ys[k + 1] - ys[k], 2) for k in range(8)]
            xs = []; last = None; n = 0; per = []
            for c in range(p1.Start, p1.End - 1):
                x = round(d.Range(c, c).Information(5), 2)
                if x != last:
                    if last is not None: per.append(n)
                    xs.append(x); last = x; n = 0
                n += 1
            per.append(n)
            x2 = round(d.Range(p2.Start, p2.Start).Information(5), 2)
            print(f'sz={sz / 2} linePitch={lp / 20} charSpace={cs} ({cs / 4096:.3f}pt) | adv={adv} | p1 cols x={xs} per_col={per} | p2 x={x2} col_pitch={round(xs[0] - x2, 2) if len(xs) == 1 else round(xs[0] - xs[1], 2)}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
