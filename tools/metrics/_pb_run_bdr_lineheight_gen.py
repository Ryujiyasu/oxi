# -*- coding: utf-8 -*-
"""How much taller is a line whose run carries a CHARACTER border (`w:bdr`)?

technical__5175ec20 p8: the paragraphs '変更前' and '変更後' sit directly above a table and
are the only two on the page whose run carries `<w:bdr w:val="single" w:sz="4" w:space="0"
w:color="auto"/>`. Word puts the table's top rule 15.61 / 15.85 below those paragraphs'
Information(6); a plain 10.5 ＭＳ 明朝 line is 13.5 and the table's own top border draws
0.48, so the character border is worth 1.63. Oxi gives 14.00 / 14.50 = 13.5 + the table
border only, i.e. it does not grow the line box at all. Two of these are the ~3pt that
keeps the 3.18 heading on page 8 where Word pushes it to page 9.

Sheet: one plain paragraph, then N=6 paragraphs whose whole run carries the border, then a
plain paragraph. The per-line height is read as (y[6]-y[1])/5 so Info6's 0.75 quantisation
averages out, and the plain control is read the same way from a second document.

Arms: border weight (sz 4 / 8 / 12 / 24) x w:space (0 / 10 / 20) x style (single / double /
dashed) x coverage (whole run / only the middle characters) x font size (21 / 24).
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/run_bdr_lh'); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>')

N = 6


def styles(sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style></w:styles>')


def document(bsz, space, style, coverage, sz):
    if bsz is None:
        rpr = ''
    else:
        rpr = f'<w:rPr><w:bdr w:val="{style}" w:sz="{bsz}" w:space="{space}" w:color="auto"/></w:rPr>'
    if coverage == 'part' and rpr:
        run = f'<w:r><w:t>変</w:t></w:r><w:r>{rpr}<w:t>更</w:t></w:r><w:r><w:t>前</w:t></w:r>'
    else:
        run = f'<w:r>{rpr}<w:t>変更前</w:t></w:r>'
    body = '<w:p><w:r><w:t>うえの段落です。</w:t></w:r></w:p>'
    body += ''.join(f'<w:p>{run}</w:p>' for _ in range(N))
    body += '<w:p><w:r><w:t>したの段落です。</w:t></w:r></w:p>'
    sect = ('<w:sectPr><w:pgSz w:w="11907" w:h="16840" w:code="9"/><w:pgMar w:top="1440" w:right="1418" w:bottom="1440" w:left="1418" w:header="720" w:footer="720"/>'
            '<w:cols w:space="425"/><w:docGrid w:linePitch="360"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = [(None, 0, 'single', 'all', 21)]
    for bsz in (4, 8, 12, 24):
        arms.append((bsz, 0, 'single', 'all', 21))
    for space in (10, 20):
        arms.append((4, space, 'single', 'all', 21))
    arms += [(4, 0, 'double', 'all', 21), (4, 0, 'dashed', 'all', 21),
             (4, 0, 'single', 'part', 21), (4, 0, 'single', 'all', 24), (None, 0, 'single', 'all', 24)]
    for bsz, space, style, coverage, sz in arms:
        tag = f'b{bsz}_s{space}_{style}_{coverage}_sz{sz}'
        at = OUT / f'{tag}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', styles(sz))
            z.writestr('word/document.xml', document(bsz, space, style, coverage, sz))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            ys = []
            for i in range(1, N + 3):
                r = d.Paragraphs(i).Range
                ys.append(round(d.Range(r.Start, r.Start).Information(6), 2))
            exact = round((ys[N] - ys[1]) / (N - 1), 4)
            print(f'bdr sz={str(bsz):4} space={space:3} style={style:7} cover={coverage:4} sz={sz} | line={exact:8} ys={ys}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
