# -*- coding: utf-8 -*-
"""How many vertical (tbRl) columns does Word fit across the text area?

creative__25b9ec89 (landscape A4, tbRl, docGrid lines 360, margins L 1701 / R 1985 = text width
657.6pt = 36.53 pitches): Word's leftmost text column on p5 sits at x=130.5 (35 columns from the
right margin at 742.65) and the next paragraph opens p6; Oxi fills a 36th column at x=94.65.

Sheet: the document's section, N one-column paragraphs ('col k' + a few chars), left margin
swept so the leftover fraction of a pitch changes; readout = number of paragraphs Word places on
page 1 (= columns per page) and the x of the last one.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/vcols'); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>')
RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr>'
N = int(os.environ.get('N', '60'))


def document(left, right, pitch):
    body = ''.join(f'<w:p><w:pPr>{RPR}</w:pPr><w:r>{RPR}<w:t>列{k:02d}の文字</w:t></w:r></w:p>' for k in range(N))
    sect = (f'<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1701" w:right="{right}" w:bottom="1701" w:left="{left}" w:header="851" w:footer="992" w:gutter="0"/>'
            f'<w:cols w:space="425"/><w:textDirection w:val="tbRl"/><w:docGrid w:type="lines" w:linePitch="{pitch}"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = [(1701, 1985, 360), (1600, 1985, 360), (1500, 1985, 360), (1800, 1985, 360), (1900, 1985, 360), (2000, 1985, 360), (1701, 1985, 400), (1701, 1985, 320)]
    for left, right, pitch in arms:
        at = OUT / f'l{left}_r{right}_p{pitch}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/document.xml', document(left, right, pitch))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            n1 = 0; xs = []
            for k in range(1, d.Paragraphs.Count + 1):
                r = d.Paragraphs(k).Range; s = d.Range(r.Start, r.Start)
                if s.Information(3) == 1:
                    n1 += 1; xs.append(round(s.Information(5), 2))
                else:
                    break
            width = (16838 - left - right) / 20.0; p = pitch / 20.0
            print(f'left={left} right={right} pitch={p} text_w={width:.2f} ({width / p:.3f} pitches) -> cols on p1 = {n1}, first x={xs[0]}, last x={xs[-1]}, left margin x={left / 20.0:.2f}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
