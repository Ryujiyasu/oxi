# -*- coding: utf-8 -*-
"""Does a table cell's line snap to the document's line grid, and what does
`<w:doNotSnapToGridInCell/>` change?

forms__008a2f3e (docGrid type="lines" linePitch=300 = 15pt, compat 11, settings carry
doNotSnapToGridInCell): row 1 is `trHeight hRule="exact" 743` (37.15pt) and holds a 16pt
'履 歴 書'. Word centres a ~20.75pt line in it (text top 87.0 for a row top of 78.3); Oxi
centres a ~28.9pt line (text top 82.4), i.e. Oxi snapped the 16pt line up to 2 grid rows.
The 4.6pt is the first step of the ~6pt page-1 drift that makes Oxi fit one line too many.

Sheet: a 3-row table (exact heights) whose first cell holds one line at SZ, on a lines-grid
page. Arms: doNotSnapToGridInCell present/absent x paragraph snapToGrid on/off x SZ.
Readout: the y of the cell line and of the row below it.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/cellsnapgrid'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')


def settings(no_snap):
    flag = '<w:doNotSnapToGridInCell/>' if no_snap else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat>{flag}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat></w:settings>')


def document(sz, snap_to_grid, row_exact, bold=False):
    b = '<w:b/><w:bCs/>' if bold else ''
    rpr = f'<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>{b}<w:kern w:val="0"/><w:sz w:val="{sz}"/></w:rPr>'
    snap = '' if snap_to_grid else '<w:snapToGrid w:val="0"/>'
    big = f'<w:p><w:pPr>{snap}{rpr}</w:pPr><w:r>{rpr}<w:t>大字</w:t></w:r></w:p>'
    small_rpr = '<w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:hint="eastAsia"/><w:sz w:val="21"/></w:rPr>'
    small = f'<w:p><w:pPr>{snap}{small_rpr}</w:pPr><w:r>{small_rpr}<w:t>小字</w:t></w:r></w:p>'
    trh = f'<w:trPr><w:trHeight w:hRule="exact" w:val="{row_exact}"/></w:trPr>' if row_exact else ''
    rows = (f'<w:tr>{trh}<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{big}</w:tc></w:tr>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{small}</w:tc></w:tr>')
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>{rows}</w:tbl>')
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="539" w:right="1531" w:bottom="539" w:left="1531" w:header="851" w:footer="992"/>'
            '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="300" w:charSpace="532"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{tbl}<w:p/>{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    if os.environ.get('BOLD_ARMS'):
        arms = [(32, True, 0, True, True), (32, True, 0, True, False), (32, True, 743, True, True), (21, True, 0, True, True)]
    else:
        arms = []
        for no_snap in (True, False):
            for sz in (32, 21):
                arms.append((sz, True, 0, no_snap, False))
        arms += [(32, False, 0, True, False), (32, False, 0, False, False), (32, True, 743, True, False), (32, True, 743, False, False)]
    for sz, snap, row_exact, no_snap, bold in arms:
        at = OUT / f'sz{sz}_snap{int(snap)}_exact{row_exact}_nosnapcell{int(no_snap)}_b{int(bold)}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', settings(no_snap)); z.writestr('word/document.xml', document(sz, snap, row_exact, bold))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            t = d.Tables(1)
            y1 = round(d.Range(t.Cell(1, 1).Range.Start, t.Cell(1, 1).Range.Start).Information(6), 2)
            y2 = round(d.Range(t.Cell(2, 1).Range.Start, t.Cell(2, 1).Range.Start).Information(6), 2)
            print(f'sz={sz / 2:5} snapToGrid={int(snap)} rowExact={row_exact:4} doNotSnapInCell={int(no_snap)} bold={int(bold)} | y_row1={y1} y_row2={y2} row1_height={round(y2 - y1, 2)}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
