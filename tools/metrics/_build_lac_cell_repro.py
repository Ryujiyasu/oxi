# -*- coding: utf-8 -*-
"""Minimal linesAndChars docGrid repro to isolate the CELL vs BODY grid line
height for 10.5pt CJK text (b35123 = tokumei_08_01 regime: linePitch=350=17.5pt,
charSpace=-2714). Measures per-line baseline gap via Word PDF glyphs vs Oxi
--dump-glyphs. Anchors bracket each block so absolute drift is separable.

Question: is Word's linesAndChars line height exactly the grid pitch (17.5pt)
for a 10.5pt line, in BODY and in a table CELL? And does charSpace change it?
"""
import os, zipfile
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
FONT = 'w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"'
# long CJK text that wraps to 5-6 lines in a ~400pt column (validate the
# per-line device-snap oscillation 13.56/13.68 across many consecutive lines)
LONG = ("業務データを適正に管理するために必要な措置とその内容について定めることとし関係法規を遵守する"
        "とともに個人の権利利益を保護することを目的として組織的人的物理的及び技術的な安全管理措置を"
        "講じるものとし取扱状況を定期的に点検し必要な改善を図ることによって適正な取扱いを確保する")

def rpr(sz=21):
    return f'<w:rPr><w:rFonts {FONT}/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'

def anchor(n):
    return f'<w:p><w:pPr>{rpr()}</w:pPr><w:r>{rpr()}<w:t>ANCHOR{n}</w:t></w:r></w:p>'

def body_para(text, sz=21):
    return f'<w:p><w:pPr>{rpr(sz)}</w:pPr><w:r>{rpr(sz)}<w:t>{text}</w:t></w:r></w:p>'

def cell_table(text, sz=21):
    # single-cell table, cell ~400pt wide (8000 dxa), so text wraps
    cellp = f'<w:p><w:pPr>{rpr(sz)}</w:pPr><w:r>{rpr(sz)}<w:t>{text}</w:t></w:r></w:p>'
    return ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '</w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>{cellp}</w:tc></w:tr>'
            '</w:tbl>')

body = []
body.append(anchor("BODY_A"))
body.append(body_para(LONG))           # body 10.5pt, wraps
body.append(anchor("BODY_B"))
body.append(anchor("CELL_A"))
body.append(cell_table(LONG))          # cell 10.5pt, wraps
body.append(anchor("CELL_B"))

# linesAndChars, linePitch=350 (17.5pt), charSpace=-2714 (b35123 exact)
sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1418" w:right="1418" w:bottom="1134" w:left="1418" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/></w:sectPr>')

doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       f'<w:document {NS}><w:body>' + ''.join(body) + sect + '</w:body></w:document>')
ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
out = os.path.join(os.path.dirname(__file__), "..", "golden-test", "repros", "lac_cell", "lac_cell.docx")
os.makedirs(os.path.dirname(out), exist_ok=True)
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", ct)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", doc)
print("wrote", os.path.abspath(out))
