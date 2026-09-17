# -*- coding: utf-8 -*-
"""What line height does Word give a BARE EMPTY paragraph on a no-type docGrid?

technical__5175ec20 p8 (pgMar bottom 1440, docGrid linePitch=360 with NO w:type, Normal =
ASCII Century / eastAsia MS Mincho 10.5): Word's five consecutive empty paragraphs step
13.5 / 14.25 alternating (mean 13.875 = an exact accumulation quantised to 0.75 per the
Info6 latin-pixel law), so five of them put the next y at 726.0 and the following heading
plus its follower overflow 770 -> Word pushes both via keepNext. Oxi steps 13.25 each,
i.e. 0.62pt short per empty = 3.1pt over five, and fits one paragraph too many. That same
-1 is all three remaining markers-off JA failures.

Sheet: one text paragraph, then N=8 bare empty paragraphs, then a text paragraph. The
per-empty height is read as (y_last_empty - y_first_empty) / (N-1) so the 0.75 Info6
quantisation averages out to the exact value.

Arms: ascii font (Century / Times New Roman / MS Mincho) x sz (21 / 24) x docGrid
(linePitch 360 no-type / linePitch 240 no-type / type=lines 360 / absent). eastAsia is
held at MS Mincho throughout, so any movement is the ASCII face's box (S195/S583).
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/emptypara_lh'); OUT.mkdir(parents=True, exist_ok=True)
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

N_EMPTY = 8


def styles(ascii_font, sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{ascii_font}" w:eastAsia="ＭＳ 明朝" w:hAnsi="{ascii_font}" w:cs="Times New Roman"/>'
            f'<w:kern w:val="2"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
            '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')


def document(grid):
    body = '<w:p><w:r><w:t>さいしょの段落です。</w:t></w:r></w:p>' + ('<w:p/>' * N_EMPTY) + '<w:p><w:r><w:t>おわりの段落です。</w:t></w:r></w:p>'
    if grid == 'none':
        g = ''
    elif grid == 'lines360':
        g = '<w:docGrid w:type="lines" w:linePitch="360"/>'
    else:
        g = f'<w:docGrid w:linePitch="{grid}"/>'
    sect = ('<w:sectPr><w:pgSz w:w="11907" w:h="16840" w:code="9"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720"/>'
            f'<w:cols w:space="425"/>{g}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = []
    for font in ('Century', 'Times New Roman', 'ＭＳ 明朝'):
        arms.append((font, 21, '360'))
    for sz in (24, 18):
        arms.append(('Century', sz, '360'))
    arms.append(('Times New Roman', 24, '360'))
    arms.append(('Times New Roman', 24, 'none'))
    for grid in ('240', 'lines360', 'none'):
        arms.append(('Century', 21, grid))
    for font, sz, grid in arms:
        tag = f'{font.replace(" ", "")}_{sz}_{grid}'
        at = OUT / f'{tag}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', styles(font, sz))
            z.writestr('word/document.xml', document(grid))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            ys = []
            for i in range(1, N_EMPTY + 3):
                r = d.Paragraphs(i).Range
                ys.append(round(d.Range(r.Start, r.Start).Information(6), 2))
            steps = [round(ys[i + 1] - ys[i], 2) for i in range(len(ys) - 1)]
            exact = round((ys[N_EMPTY] - ys[1]) / (N_EMPTY - 1), 4)
            print(f'ascii={font:16} sz={sz:3} grid={grid:9} | empty_exact={exact:8} steps={steps}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
