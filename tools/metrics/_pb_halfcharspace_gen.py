# -*- coding: utf-8 -*-
"""On a `docGrid type="linesAndChars"` page, how much of `charSpace` does a HALF-WIDTH
character carry?

technical__9e4d04b4 (charSpace -3531 = -0.862pt, Century 10.5): Word advances 'm' 9.0, '(' 3.0,
'G' 7.5, 'y' 5.25 where Century's own widths are 9.336 / 3.497 / 8.167 / 5.640 — each is
(natural - 0.431) quantised to 0.75, i.e. HALF the charSpace. A full-width character carries the
whole charSpace (10.5 - 0.862 = 9.638), and an eastAsia-hinted Greek mu is full-width (9.75).
Oxi gives the Latin characters their natural width, so a cell line measures 2.3pt wide and wraps.

Sheet: one paragraph 'mGy(x)' + 6 kana per arm, charSpace swept over negative / zero / positive,
Latin font Century / Times New Roman. Readout: per-character advances (Information 5 deltas).
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/halfcharspace'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
TEXT = os.environ.get('TEXT', 'mGy(x)μあいうえ')


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat>{BAL}{FE}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


BAL = '<w:balanceSingleByteDoubleByteWidth/>' if os.environ.get('BAL') else ''
FE = '<w:useFELayout/>' if os.environ.get('FE') else ''


def document(latin, char_space, grid_type, in_cell=False, kern=0):
    kx = f'<w:kern w:val="{kern}"/>' if kern else ''
    rpr = f'<w:rPr><w:rFonts w:ascii="{latin}" w:eastAsia="ＭＳ 明朝" w:hAnsi="{latin}" w:hint="eastAsia"/>{kx}</w:rPr>'
    jc = os.environ.get('JC', '')
    jx = f'<w:jc w:val="{jc}"/>' if jc else ''
    fc = os.environ.get('FC', '')
    fx = f'<w:ind w:firstLineChars="{fc}" w:firstLine="{int(round(int(fc)/100.0*192.7588))}"/>' if fc else ''
    para = f'<w:p><w:pPr>{fx}{jx}{rpr}</w:pPr><w:r>{rpr}<w:t>{TEXT}</w:t></w:r></w:p>'
    if in_cell:
        body = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
                '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
                f'<w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>{para}</w:tc></w:tr></w:tbl><w:p/>')
    else:
        body = para
    grid = '' if grid_type == 'none' else f'<w:docGrid w:type="{grid_type}" w:linePitch="291" w:charSpace="{char_space}"/>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1418" w:right="851" w:bottom="1134" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>'
            f'<w:cols w:space="425"/>{grid}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    if os.environ.get('KERN_ARMS'):
        arms = [('Century', -3531, 'linesAndChars', 11, False, k) for k in (0, 2, 20)]
        arms += [('Century', 0, 'linesAndChars', 11, False, 2), ('Century', -3531, 'linesAndChars', 11, True, 2)]
    elif os.environ.get('CELL_ARMS'):
        arms = [(latin, cs, 'linesAndChars', 11, True, 0) for latin in ('Century',) for cs in (-3531, -2048, 0)]
        arms += [('Century', -3531, 'linesAndChars', 11, False, 0), ('Times New Roman', -3531, 'linesAndChars', 11, True, 0)]
    else:
        arms = [(latin, cs, 'linesAndChars', 11, False, 0) for latin in ('Century', 'Times New Roman') for cs in (-3531, -2048, 0, 2048)]
        arms += [('Century', -3531, 'none', 11, False, 0), ('Century', -3531, 'lines', 11, False, 0), ('Century', -3531, 'linesAndChars', 15, False, 0)]
    for latin, cs, gt, compat, in_cell, kern in arms:
        at = OUT / f'{latin.replace(" ", "")}_cs{cs}_{gt}_c{compat}_{"cell" if in_cell else "body"}_k{kern}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', settings(compat)); z.writestr('word/document.xml', document(latin, cs, gt, in_cell, kern))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            p = (d.Tables(1).Cell(1, 1).Range.Paragraphs(1).Range if in_cell else d.Paragraphs(1).Range)
            xs = [round(d.Range(c, c).Information(5), 2) for c in range(p.Start, p.End - 1)]
            adv = [round(xs[k + 1] - xs[k], 2) for k in range(len(xs) - 1)]
            print(f'{latin:16} cs={cs:6} ({cs / 4096.0:+.3f}) {gt:13} c{compat} {"cell" if in_cell else "body":4} kern={kern:2} span6={round(xs[6] - xs[0], 2)} adv={adv}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
