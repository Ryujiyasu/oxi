# -*- coding: utf-8 -*-
"""What space does Word put before a paragraph that inherits `w:before` from its STYLE and
carries `w:beforeLines` directly?

technical__5175ec20 p8: '3.18 …' is pStyle "2" (heading 2, style `<w:spacing w:before="240"
w:after="60"/>` = 12pt) with a DIRECT `<w:spacing w:beforeLines="100"/>` and no direct w:before.
Whether the space is 12pt (the style's before / the earlier probe's "unit is 12pt") or 18pt (the
section's docGrid linePitch 360) decides whether the heading and its follower still fit above the
page bottom at 770 — Word pushes them, so 18 is the candidate. Oxi emits 9.65.

Sheet: text paragraph, then the heading paragraph, then a follower. Arms: style before
(240 / 0 / absent) x direct beforeLines (0 / 50 / 100 / 200) x linePitch (360 / 240 / none).
Readout: y of all three paragraphs, so "space before" = y2 - (y1 + line1).
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/beforelines_style'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def styles(style_before, dd_sz=21, normal_sz=None, h2_sz=24):
    """dd_sz = docDefaults size; normal_sz = the Normal style's own size (None = inherit).

    S1453 の導出では dd_sz=21 / normal_sz=None の 1 点しか測っておらず、そこから
    「beforeLines の単位は 12pt 固定」と一般化してしまった。policies__094c44cd
    (docDefaults に sz 無し、Normal が sz=24) は beforeLines=50 に 7.52pt を与え、
    12pt 単位なら 6.0 のはずのところと合わない。単位が既定サイズに依存するかを
    決めるため、docDefaults と Normal のサイズを振る。
    """
    sb = '' if style_before is None else f'<w:spacing w:before="{style_before}" w:after="60"/>'
    dd = f'<w:sz w:val="{dd_sz}"/>' if dd_sz else ''
    nrm = f'<w:rPr><w:sz w:val="{normal_sz}"/></w:rPr>' if normal_sz else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
            f'{dd}</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            f'<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>{nrm}</w:style>'
            f'<w:style w:type="paragraph" w:styleId="H2"><w:name w:val="heading 2"/><w:basedOn w:val="a"/><w:pPr><w:keepNext/>{sb}<w:outlineLvl w:val="1"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="ＭＳ ゴシック" w:hAnsi="Arial"/><w:b/><w:sz w:val="{h2_sz}"/></w:rPr></w:style></w:styles>')


def document(before_lines, pitch):
    bl = f'<w:spacing w:beforeLines="{before_lines}"/>' if before_lines else ''
    body = ('<w:p><w:r><w:t>まえの段落です。</w:t></w:r></w:p>'
            f'<w:p><w:pPr><w:pStyle w:val="H2"/>{bl}</w:pPr><w:r><w:t>3.18 見出し</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>あとの段落です。</w:t></w:r></w:p>')
    grid = f'<w:docGrid w:linePitch="{pitch}"/>' if pitch else ''
    sect = ('<w:sectPr><w:pgSz w:w="11907" w:h="16840" w:code="9"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720"/>'
            f'<w:cols w:space="425"/>{grid}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = []
    for sb in (240, 0, None):
        for bl in (0, 100):
            arms.append((sb, bl, 360))
    arms += [(240, 50, 360), (240, 200, 360), (240, 100, 240), (240, 100, None), (None, 100, 240)]
    # size-varying arms: (dd_sz, normal_sz, h2_sz) with a fixed beforeLines=50
    size_arms = [
        (21, None, 24), (21, None, 21), (None, 24, 24), (None, 24, 21),
        (24, None, 24), (None, 21, 24), (None, None, 24),
    ]
    for dd, nz, hz in size_arms:
        tag = f'dd{dd}_n{nz}_h{hz}'
        at = OUT / f'size_{tag}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', styles(240, dd, nz, hz))
            z.writestr('word/document.xml', document(50, 360))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            ys = [round(d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start).Information(6), 2) for i in (1, 2, 3)]
            print(f'SIZE dd={str(dd):4} normal={str(nz):4} h2={hz} | y={ys} gap1={round(ys[1] - ys[0], 2)} gap2={round(ys[2] - ys[1], 2)}', flush=True)
        finally:
            d.Close(False)
    for sb, bl, pitch in arms:
        at = OUT / f'sb{sb}_bl{bl}_p{pitch}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', styles(sb))
            z.writestr('word/document.xml', document(bl, pitch))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            ys = [round(d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start).Information(6), 2) for i in (1, 2, 3)]
            line1 = ys[1] - ys[0]
            line2 = ys[2] - ys[1]
            print(f'styleBefore={str(sb):5} beforeLines={bl:4} linePitch={str(pitch):5} | y={ys} gap1={round(line1, 2)} gap2={round(line2, 2)}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
