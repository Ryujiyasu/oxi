# -*- coding: utf-8 -*-
"""On a `docGrid` with NO type attribute, does Word advance each line by the floor-to-0.75
line height, or by the exact height with only the painted position snapped?

legal__05c84880 (游明朝 10.5, docGrid linePitch=286 no type): Word's line tops read
17.25 / 18.0 / 17.25 ... = the exact 17.527 (Yu Mincho win-sum 1.2871 x 10.5 x 83/64) accumulated
with each position snapped to the 96dpi pixel. Oxi floors the ADVANCE to 17.25 and loses
0.277pt per line, which is the ~28pt drift per page behind its lastRenderedPageBreak dependence.

Sheet: one paragraph of N lines per arm (font x size x linePitch), readout = the y of every line
so both the per-line deltas and the cumulative position are visible.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/notypegrid'); OUT.mkdir(parents=True, exist_ok=True)
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
LINE = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれ'


def document(font, sz, line_pitch, nlines):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:hint="eastAsia"/><w:sz w:val="{sz}"/></w:rPr>'
    body = ''.join(f'<w:p><w:pPr><w:jc w:val="both"/>{rpr}</w:pPr><w:r>{rpr}<w:t>{LINE}</w:t></w:r></w:p>' for _ in range(nlines))
    grid = f'<w:docGrid w:linePitch="{line_pitch}"/>' if line_pitch else ''
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/>'
            f'<w:cols w:space="425"/>{grid}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = [('游明朝', 21, 286), ('游明朝', 24, 286), ('游ゴシック', 21, 286),
            ('ＭＳ 明朝', 21, 286), ('メイリオ', 21, 286), ('游明朝', 21, 360), ('游明朝', 21, 0)]
    for font, sz, lp in arms:
        at = OUT / f'{font.replace(" ", "")}_{sz}_lp{lp}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/document.xml', document(font, sz, lp, 8))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            ys = []
            for i in range(1, 9):
                r = d.Paragraphs(i).Range
                ys.append(round(d.Range(r.Start, r.Start).Information(6), 2))
            adv = [round(ys[k + 1] - ys[k], 2) for k in range(len(ys) - 1)]
            mean = round((ys[-1] - ys[0]) / (len(ys) - 1), 4)
            print(f'{font:8} {sz / 2:5} linePitch={lp / 20 if lp else 0:6} | y0={ys[0]} adv={adv} mean={mean}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
