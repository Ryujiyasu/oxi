# -*- coding: utf-8 -*-
"""In vertical text (tbRl), does Word advance a proportional CJK font's characters by their
proportional width instead of the em?

creative__25b9ec89 (ＭＳ Ｐ明朝 11pt, landscape, tbRl): Word's per-character y advances are
9.75 / 10.5 / 7.5 (、) / 9.75 / 8.25 (ン) — the horizontal proportional widths (S1415) — so a
column holds 41-44 characters; Oxi advances 11pt each and packs a different count.

Sheet: vertical section, one paragraph 'ふと、トンネルの奥から踏切でカンカン鳴る音が聞こえた。' per
font × size. Readout: per-character Information(6) advances for the first 12 characters.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/vprop'); OUT.mkdir(parents=True, exist_ok=True)
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
TEXT = 'ふと、トンネルの奥から踏切でカンカン鳴る音が聞こえた。'


def document(font, sz):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:hint="eastAsia"/><w:kern w:val="0"/><w:sz w:val="{sz}"/></w:rPr>'
    body = f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>{TEXT}</w:t></w:r></w:p>'
    sect = ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1701" w:right="1985" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="425"/><w:textDirection w:val="tbRl"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for font in ('ＭＳ 明朝', 'ＭＳ Ｐ明朝', 'ＭＳ Ｐゴシック', 'ＭＳ ゴシック', '游明朝', 'メイリオ'):
        for sz in (21, 22, 24):
            at = OUT / f'{font.replace(" ", "")}_{sz}.docx'
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
                z.writestr('word/settings.xml', SETTINGS); z.writestr('word/document.xml', document(font, sz))
            d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
            try:
                p = d.Paragraphs(1).Range
                ys = [round(d.Range(c, c).Information(6), 2) for c in range(p.Start, p.Start + 13)]
                adv = [round(ys[k + 1] - ys[k], 2) for k in range(12)]
                print(f'{font:10} {sz / 2:5} name={p.Font.NameFarEast} adv={adv}', flush=True)
            finally:
                d.Close(False)
finally:
    app.Quit()
