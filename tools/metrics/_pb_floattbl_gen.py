# -*- coding: utf-8 -*-
"""How does Word lay out a page-anchored, full-width floating table when the body flow
(the continuation of a split inline table) also lands on that page?

policies__00602e8a p4 (Word): floating table 4 (tblpPr vertAnchor=page tblpY=2431) shows
ONLY its row 0 at 122pt; the previous inline table's continuation follows at 366.75 (row 0
bottom + 9pt); the float's rows 1-3 appear on p5 at 114.75. Oxi keeps rows 1-3 on p4.

Sheet (A4 portrait, 1in margins): 'before', inline table A (row 0 one line, row 1 = LINES_A
lines so that it splits onto p2 with CONT lines), paragraph 'mid', floating table B
(vertAnchor=page, tblpY = TBLPY twips; rows: r0 = LINES_R0 lines, r1 'header' 1 line,
r2 = LINES_R2 lines, r3 1 line), 'after'. Readout per arm: page/y of the first char of A row 1's
continuation (first char after the split), of B rows 0-3, and of 'after'.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/floattbl'); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>'
LINE = 'line text line text line text line text line text line text line text line text end'  # ~1 full line at 11pt/468pt


def para(text):
    return f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{RPR}</w:pPr><w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def cell(paras, w=9360):
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{"".join(paras)}</w:tc>'


def lines_paras(prefix, n):
    return [para(f'{prefix}{k:02d} ' + LINE) for k in range(n)]


TBLPR_INLINE = ('<w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
                '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>')


def tblpr_float(tblpy):
    return TBLPR_INLINE.replace('<w:tblPr>', f'<w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="page" w:horzAnchor="margin" w:tblpY="{tblpy}"/>')


def document(lines_a, tblpy, lines_r0, lines_r2):
    ta = f'<w:tbl>{TBLPR_INLINE}<w:tr>{cell([para("A row0")])}</w:tr><w:tr>{cell(lines_paras("A1-", lines_a))}</w:tr></w:tbl>'
    tb = (f'<w:tbl>{tblpr_float(tblpy)}<w:tr>{cell(lines_paras("B0-", lines_r0))}</w:tr><w:tr>{cell([para("B1 header")])}</w:tr>'
          f'<w:tr>{"<w:trPr><w:trHeight w:val=" + chr(34) + os.environ["TRH2"] + chr(34) + "/></w:trPr>" if os.environ.get("TRH2") else ""}{cell(lines_paras("B2-", lines_r2))}</w:tr><w:tr>{cell([para("B3 last")])}</w:tr></w:tbl>')
    body = para('before') + ta + para('mid') + tb + para('after')
    sect = ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' if os.environ.get('LANDSCAPE') else '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>') + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


def pos(d, rng):
    r = d.Range(rng.Start, rng.Start)
    return (r.Information(3), round(r.Information(6), 2))


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = []
    for lines_a in (int(x) for x in os.environ.get('LINES_A', '48,55').split(',')):
        for tblpy in (int(x) for x in os.environ.get('TBLPY', '2160,4320').split(',')):
            for lines_r0 in (int(x) for x in os.environ.get('LINES_R0', '3,12').split(',')):
                for lines_r2 in (int(x) for x in os.environ.get('LINES_R2', '5').split(',')):
                    arms.append((lines_a, tblpy, lines_r0, lines_r2))
    for lines_a, tblpy, lines_r0, lines_r2 in arms:
        at = OUT / f'a{lines_a}_y{tblpy}_r0{lines_r0}_r2{lines_r2}{"_land" if os.environ.get("LANDSCAPE") else ""}{"_trh" + os.environ["TRH2"] if os.environ.get("TRH2") else ""}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/document.xml', document(lines_a, tblpy, lines_r0, lines_r2))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            ta = d.Tables(1); tb = d.Tables(2)
            # A row1: first paragraph on the page after the split
            a_paras = ta.Cell(2, 1).Range.Paragraphs
            first_pg = pos(d, a_paras(1).Range)[0]
            cont = None
            for k in range(1, a_paras.Count + 1):
                p = pos(d, a_paras(k).Range)
                if p[0] != first_pg:
                    cont = (k, p); break
            rows = [pos(d, tb.Cell(r, 1).Range) for r in (1, 2, 3, 4)]
            after = pos(d, d.Paragraphs(d.Paragraphs.Count).Range)
            mid = None
            for k in range(1, d.Paragraphs.Count + 1):
                if d.Paragraphs(k).Range.Text.startswith('mid'):
                    mid = pos(d, d.Paragraphs(k).Range); break
            print(f'A={lines_a} tblpY={tblpy} r0={lines_r0} r2={lines_r2} | A-row1 starts p{first_pg} cont={cont} mid={mid} | B rows={rows} | after={after} pages={d.ComputeStatistics(2)}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
