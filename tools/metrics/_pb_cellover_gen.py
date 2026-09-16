# -*- coding: utf-8 -*-
"""How far past a table cell's right text edge may the last character sit before Word wraps it?

technical__9e4d04b4 row 2 (compat 11, ＭＳ 明朝 10.5, cell 183.4pt wide, tcMar 5.4pt each side =
172.6pt of text): a right-aligned paragraph of 8 leading ideographic spaces + 'm(μ)Gy／時間　以下'
measures 175.3pt, so the last character ends 1.9pt past the text edge — Word keeps it on one
line, Oxi wraps it.

Sheet: a one-column table, cell width swept in 1.5pt steps around the text width of
'あいうえおかきくけこさしすせそたちつてと' (20 chars at 10.5 = 202.4 on this doc's grid), with
tcMar and paragraph alignment as arms. Readout: characters on line 1 of the cell paragraph.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/cellover'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
KANA = 'あいうえおかきくけこさしすせそたちつてと'


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


def document(cell_tw, mar_tw, jc, n, grid):
    rpr = '<w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:hint="eastAsia"/></w:rPr>'
    jc_xml = f'<w:jc w:val="{jc}"/>' if jc != 'left' else ''
    p = f'<w:p><w:pPr>{jc_xml}{rpr}</w:pPr><w:r>{rpr}<w:t>{KANA[:n]}</w:t></w:r></w:p>'
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           f'<w:tblCellMar><w:left w:w="{mar_tw}" w:type="dxa"/><w:right w:w="{mar_tw}" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="{cell_tw}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="{cell_tw}" w:type="dxa"/></w:tcPr>{p}</w:tc></w:tr></w:tbl>')
    g = '<w:docGrid w:type="linesAndChars" w:linePitch="291" w:charSpace="-3531"/>' if grid else ''
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1418" w:right="851" w:bottom="1134" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>'
            f'<w:cols w:space="425"/>{g}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{tbl}<w:p/>{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    # text width of 20 kana at 10.5 on this grid = 20 * 9.638 = 192.76pt = 3855tw
    arms = []
    for extra in (-60, -40, -20, -10, 0, 10, 20, 40):
        arms.append((3855 + 216 + extra, 108, 'left', 20, True))
    arms += [(3855 + 216, 108, 'right', 20, True), (3855 + 216, 108, 'center', 20, True),
             (3855 + 216 - 20, 108, 'right', 20, True), (3855 + 216 - 20, 0, 'left', 20, True),
             (3855 + 216 - 20, 108, 'left', 20, False)]
    for cell_tw, mar_tw, jc, n, grid in arms:
        at = OUT / f'w{cell_tw}_m{mar_tw}_{jc}_n{n}_{"grid" if grid else "nogrid"}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', settings(11)); z.writestr('word/document.xml', document(cell_tw, mar_tw, jc, n, grid))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            cell = d.Tables(1).Cell(1, 1)
            p = cell.Range.Paragraphs(1).Range
            rows = [(round(d.Range(c, c).Information(5), 2), round(d.Range(c, c).Information(6), 2)) for c in range(p.Start, p.End - 1)]
            y0 = rows[0][1]
            line1 = [r for r in rows if r[1] == y0]
            inner = (cell_tw - 2 * mar_tw) / 20.0
            print(f'cellW={cell_tw / 20.0:7.2f} inner={inner:7.2f} mar={mar_tw / 20.0:5.2f} {jc:6} {"grid" if grid else "nogrid":6} | chars_line1={len(line1):3} first_x={line1[0][0]} last_x={line1[-1][0]} lines={len(set(r[1] for r in rows))}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
