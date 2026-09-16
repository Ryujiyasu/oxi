# -*- coding: utf-8 -*-
"""Does Word accumulate an `exact` trHeight exactly, or quantised to the 96dpi pixel?

forms__008a2f3e: ~13 rows of `trHeight hRule="exact" 539` (26.95pt). Oxi stacks 26.95 each;
Word's text positions suggest ~+0.14 more per row, which is the ~1.8pt page-1 drift that makes
Oxi fit one extra line. A single row's height cannot answer this — a cell paragraph's
Information(6) is itself snapped to 0.75, so any ONE delta looks like a 0.75 multiple. Stacking
N rows separates the two models: exact gives y0 + N x 26.95, per-row snapping gives y0 + N x 27.0.

Sheet: N rows of one 10.5pt line each, exact height swept; readout = the y of row 1 and row N
and the implied per-row step.
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/exactrowaccum'); OUT.mkdir(parents=True, exist_ok=True)
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


def document(n_rows, exact_tw, rule):
    rpr = '<w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:hint="eastAsia"/><w:sz w:val="21"/></w:rPr>'
    trh = f'<w:trPr><w:trHeight w:hRule="{rule}" w:val="{exact_tw}"/></w:trPr>' if rule else f'<w:trPr><w:trHeight w:val="{exact_tw}"/></w:trPr>'
    rows = ''.join(
        f'<w:tr>{trh}<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>行{k:02d}</w:t></w:r></w:p></w:tc></w:tr>'
        for k in range(n_rows))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>{rows}</w:tbl>')
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="539" w:right="1531" w:bottom="539" w:left="1531" w:header="851" w:footer="992"/>'
            '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="300" w:charSpace="532"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{tbl}<w:p/>{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = [(20, 539, 'exact'), (20, 456, 'exact'), (20, 306, 'exact'), (20, 743, 'exact'),
            (20, 539, 'atLeast'), (20, 539, None)]
    for n, tw, rule in arms:
        at = OUT / f'n{n}_h{tw}_{rule or "none"}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/document.xml', document(n, tw, rule))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            t = d.Tables(1)
            ys = []
            for i in (1, n):
                c = t.Cell(i, 1)
                ys.append(round(d.Range(c.Range.Start, c.Range.Start).Information(6), 2))
            step = (ys[1] - ys[0]) / (n - 1)
            print(f'n={n} trHeight={tw} ({tw / 20.0}pt) rule={rule} | y1={ys[0]} y{n}={ys[1]} step={step:.4f} exact={tw / 20.0} snapped={round(tw / 20.0 / 0.75) * 0.75}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
