# -*- coding: utf-8 -*-
"""How much vertical space sits between two tables separated by one bare empty paragraph?

forms__008a2f3e (docGrid type="lines" linePitch=300 = 15pt, Normal = ＭＳ 明朝 10.5): Word's PDF
rule lines put table 1's bottom at 175.58 and table 2's top at 192.02 (16.44 apart); Oxi's dump
borders read 175.65 and 190.65 (15.00). Everything before and after tracks within 0.15, so this
single 1.44pt is the whole page-1 drift that makes Oxi fit one extra line.

Sheet: table A (one exact row), one bare empty paragraph, table B (one exact row). Arms: border
weight (sz 4 / 12), the empty paragraph present/absent, and the grid on/off. Readout: the PDF
rule ys, so the gap is measured the same way on both sides.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/tblgap'); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>'
          '<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style></w:styles>')


def table(sz, label):
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
            f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>'
            f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/></w:tblBorders>'
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
            '<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="306"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r><w:t>{label}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')


def document(sz, with_para, grid):
    gap = '<w:p/>' if with_para else ''
    g = '<w:docGrid w:type="lines" w:linePitch="300" w:charSpace="532"/>' if grid else ''
    body = table(sz, 'A') + gap + table(sz, 'B') + '<w:p/>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="539" w:right="1531" w:bottom="539" w:left="1531" w:header="851" w:footer="992"/>'
            f'<w:cols w:space="425"/>{g}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client, pymupdf
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for sz in (12, 4):
        for with_para in (True, False):
            for grid in (True, False):
                at = OUT / f'sz{sz}_p{int(with_para)}_g{int(grid)}.docx'
                with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
                    z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', STYLES)
                    z.writestr('word/document.xml', document(sz, with_para, grid))
                d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
                try:
                    pdf = str((OUT / f'sz{sz}_p{int(with_para)}_g{int(grid)}.pdf').resolve())
                    d.ExportAsFixedFormat(pdf, 17)
                finally:
                    d.Close(False)
                doc = pymupdf.open(pdf); p = doc[0]
                ys = set()
                for dr in p.get_drawings():
                    for item in dr['items']:
                        if item[0] == 're' and item[1].height <= 2.0 and item[1].width > 50:
                            ys.add(round(item[1].y0, 2))
                        elif item[0] == 'l' and abs(item[1].y - item[2].y) < 0.5 and abs(item[1].x - item[2].x) > 50:
                            ys.add(round(item[1].y, 2))
                doc.close()
                r = sorted(ys)
                gap_val = round(r[2] - r[1], 2) if len(r) >= 3 else None
                print(f'borderSz={sz:3} emptyPara={int(with_para)} grid={int(grid)} | rules={r[:4]} gap(A.bottom->B.top)={gap_val}', flush=True)
finally:
    app.Quit()
