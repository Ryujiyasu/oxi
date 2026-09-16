# -*- coding: utf-8 -*-
"""What is the unit of w:spacing beforeLines when the section's docGrid has NO type
(only linePitch)?  technical__5175ec20 p8: a heading with beforeLines=100 lands ~6pt
higher in Oxi (12.1pt = font based) than Word; if Word takes the docGrid linePitch
(360 twips = 18pt) the heading plus its follower no longer fit the page and move.

Arms: linePitch 360 / 411 / no docGrid  x  font 10.5 / 12pt  x  beforeLines 50 / 100 / 200.
Readout (Word COM): y of the paragraph before and after -> gap - line height.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/beforelines'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')


def styles(sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman"/>'
            f'<w:sz w:val="{sz}"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>')


def document(pitch, bl):
    grid = f'<w:docGrid w:linePitch="{pitch}"/>' if pitch else ''
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            '<w:p><w:r><w:t>前の段落です。</w:t></w:r></w:p>'
            f'<w:p><w:pPr><w:spacing w:beforeLines="{bl}"/></w:pPr><w:r><w:t>beforeLines の段落です。</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>次の段落です。</w:t></w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720"/>'
            f'{grid}</w:sectPr></w:body></w:document>')


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for pitch in (360, 411, None):
        for sz in (21, 24):
            for bl in (50, 100, 200):
                name = f'p{pitch or 0}_sz{sz}_bl{bl}'
                at = os.path.join(OUT, name + '.docx')
                with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
                    z.writestr('word/styles.xml', styles(sz)); z.writestr('word/document.xml', document(pitch, bl))
                d = app.Documents.Open(os.path.abspath(at), ReadOnly=True)
                try:
                    ys = [d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start).Information(6) for i in (1, 2, 3)]
                    line = ys[2] - ys[1]
                    print(f'{name:16} y={ys[0]:.2f}/{ys[1]:.2f}/{ys[2]:.2f} line={line:.2f} before={ys[1] - ys[0] - line:.2f} (per 100 lines: {(ys[1] - ys[0] - line) * 100 / bl:.2f})', flush=True)
                finally:
                    d.Close(False)
finally:
    app.Quit()
