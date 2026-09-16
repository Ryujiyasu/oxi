# -*- coding: utf-8 -*-
"""When does Word DROP a `firstLineChars` indent inside a table cell?

technical__9e4d04b4 table 3 row 3 col 2 (cell 182.85pt, tcMar 4.95 -> 172.95pt of text): the
paragraph carries `<w:ind w:firstLineChars="795" w:firstLine="1532"/>` (76.6pt) and Word draws
'm(μ)Gy／時間　以下' (96.0pt) starting at the cell's LEFT TEXT EDGE — the indent is not applied.
Oxi applies it, the line overflows, and the last character wraps (+14.5pt for the page).

Sheet: a one-column table (cell 202.55pt, tcMar 5.4 -> 191.75pt of text) holding one paragraph
of N kana with firstLineChars swept. Readout: x of the first character, characters on line 1,
and the line count — so "indent applied?" and "wrapped?" are both visible.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/cellfirstline'); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat></w:settings>')
KANA = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも'
CELL_TW = 4051  # 202.55pt; tcMar 108 each -> 191.75pt of text = 19.9 cells of 9.638


def document(n, first_chars, in_cell):
    rpr = '<w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:hint="eastAsia"/></w:rPr>'
    first_tw = int(round(first_chars / 100.0 * 192.7588))
    ind = '' if first_chars == 0 else f'<w:ind w:firstLineChars="{first_chars}" w:firstLine="{first_tw}"/>'
    p = f'<w:p><w:pPr>{ind}{rpr}</w:pPr><w:r>{rpr}<w:t>{KANA[:n]}</w:t></w:r></w:p>'
    if in_cell:
        body = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
                '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                f'<w:tblGrid><w:gridCol w:w="{CELL_TW}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="{CELL_TW}" w:type="dxa"/></w:tcPr>{p}</w:tc></w:tr></w:tbl><w:p/>')
    else:
        body = p
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1418" w:right="851" w:bottom="1134" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="425"/><w:docGrid w:type="linesAndChars" w:linePitch="291" w:charSpace="-3531"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = [(n, fc, True) for fc in (0, 200, 400, 600, 795, 1000, 1600) for n in (10, 16)]
    arms += [(10, 795, False), (16, 795, False)]
    for n, fc, in_cell in arms:
        at = OUT / f'n{n}_fc{fc}_{"cell" if in_cell else "body"}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/document.xml', document(n, fc, in_cell))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            p = (d.Tables(1).Cell(1, 1).Range.Paragraphs(1).Range if in_cell else d.Paragraphs(1).Range)
            end = p.End - (1 if in_cell else 1)
            rows = [(round(d.Range(c, c).Information(5), 2), round(d.Range(c, c).Information(6), 2)) for c in range(p.Start, end)]
            y0 = rows[0][1]
            line1 = [r for r in rows if r[1] == y0]
            print(f'n={n:3} firstLineChars={fc:5} {"cell" if in_cell else "body":4} | first_x={line1[0][0]:7} chars_line1={len(line1):3} lines={len(set(r[1] for r in rows))} indent_pt={round(fc / 100.0 * 9.638, 2)}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
