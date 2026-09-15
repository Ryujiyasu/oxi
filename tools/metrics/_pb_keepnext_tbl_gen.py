# -*- coding: utf-8 -*-
"""Does a keepNext heading follow a TABLE whose first row does not fit the page?

policies__094c44cd5dce58a8: 「例示と好ましい選択肢」 (`<w:keepNext/>`, next block
a table) sits at Word p14 with its table while Oxi leaves it at the p13 bottom;
policies__07543a6b9776a1cf is the same class. Oxi's keepNext look-ahead only
knows a paragraph follower.

Sheet: A4, margins 1134, ＭＳ 明朝 10.5 on an 18pt line grid, N filler
paragraphs (one line each), then the heading (keepNext on/off), then a 3-row
table whose rows are 2 lines each (cantSplit). Sweep N so the heading walks
toward the page bottom. Readout: Information(3) page of the heading and of the
table's first cell, plus the heading's Information(6) y.

  python _pb_keepnext_tbl_gen.py            # all arms
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/keepnext_tbl'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:kern w:val="2"/><w:sz w:val="21"/></w:rPr>'


def para(text, keep_next=False, split=False):
    kn = '<w:keepNext/>' if keep_next else ''
    return f'<w:p><w:pPr>{kn}<w:widowControl w:val="0"/>{RPR}</w:pPr><w:r>{RPR}<w:t>{text}</w:t></w:r></w:p>'


def table(rows=3, cant_split=True):
    cs = '<w:cantSplit/>' if cant_split else ''
    r = ''
    for i in range(rows):
        cells = ''.join(f'<w:tc><w:tcPr><w:tcW w:w="4819" w:type="dxa"/></w:tcPr><w:p><w:pPr>{RPR}</w:pPr><w:r>{RPR}<w:t>セル{i + 1}行目の一行目</w:t></w:r></w:p><w:p><w:pPr>{RPR}</w:pPr><w:r>{RPR}<w:t>二行目</w:t></w:r></w:p></w:tc>' for _ in range(2))
        r += f'<w:tr><w:trPr>{cs}</w:trPr>{cells}</w:tr>'
    return ('<w:tbl><w:tblPr><w:tblW w:w="9638" w:type="dxa"/><w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            f'</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4819"/><w:gridCol w:w="4819"/></w:tblGrid>{r}</w:tbl>')


def document(n, keep_next):
    body = ''.join(para(f'埋め草の行 {i + 1}') for i in range(n)) + para('見出し（表の前）', keep_next) + table() + para('表のあとの段落')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{SECT}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    # content height 841.9 - 113.4*2 = 615.1 -> 34 lines of 18; walk the heading over the last rows
    for keep_next in (True, False):
        for n in range(35, 41):
            at = OUT / f'{"kn" if keep_next else "nokn"}_{n:02d}.docx'
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/document.xml', document(n, keep_next)); z.writestr('word/settings.xml', SETTINGS); z.writestr('word/_rels/document.xml.rels', DR)
            d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
            try:
                h = d.Paragraphs(n + 1).Range; hc = d.Range(h.Start, h.Start)
                t = d.Tables(1); c1 = t.Cell(1, 1).Range; cc = d.Range(c1.Start, c1.Start)
                print(f'{"keepNext" if keep_next else "plain   "} N={n:2d} heading p{int(hc.Information(3))} y={hc.Information(6):.2f}  row1 p{int(cc.Information(3))} y={cc.Information(6):.2f}', flush=True)
            finally:
                d.Close(False)
finally:
    app.Quit()
