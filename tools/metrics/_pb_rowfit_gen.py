# -*- coding: utf-8 -*-
"""When does Word push a table row off the page bottom, and when does it split it?

The S1420 probe left a residual: a 2-line cantSplit row whose top sits at
759.75 on a 785.2 bottom (25.45pt left, row 27.75) goes to the next page in
Word while Oxi keeps it (its row is 27.3 and it tolerates the 1.8pt overflow).
Four near-pass documents (golden proberuby / probexpbdr, EN legal__0021d29f /
educational__003299ba, JA reports__1c313df3) all slip on a row at a page bottom.

Sheet: A4, margins 1134 (content bottom 785.2), MS Mincho 10.5, 18pt line grid,
34 filler lines, then ONE spacer paragraph with an EXACT line height X (0..30pt
in 1pt steps), then a 3-row table (rows = 2 cell lines each), cantSplit on/off.
Readout: page / y of the table's row-1 first line and its second paragraph
(line 2), and of row 2 -- whether the row moved whole, split, or stayed.

  python _pb_rowfit_gen.py [cant|split]
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/rowfit'); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'
RPR = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:kern w:val="2"/><w:sz w:val="21"/></w:rPr>'
import os
CELL_SZ = int(os.environ.get('ROWFIT_SZ', '21'))
CRPR = RPR.replace('w:val="21"', 'w:val="%d"' % CELL_SZ)


def para(text, exact=None):
    sp = f'<w:spacing w:line="{int(exact * 20)}" w:lineRule="exact"/>' if exact else ''
    return f'<w:p><w:pPr>{sp}<w:widowControl w:val="0"/>{RPR}</w:pPr><w:r>{RPR}<w:t>{text}</w:t></w:r></w:p>'


def table(rows=3, cant_split=True):
    cs = '<w:cantSplit/>' if cant_split else ''
    r = ''
    for i in range(rows):
        cells = ''.join(f'<w:tc><w:tcPr><w:tcW w:w="4819" w:type="dxa"/></w:tcPr><w:p><w:pPr>{CRPR}</w:pPr><w:r>{CRPR}<w:t>セル{i + 1}行目の一行目</w:t></w:r></w:p><w:p><w:pPr>{CRPR}</w:pPr><w:r>{CRPR}<w:t>二行目</w:t></w:r></w:p></w:tc>' for _ in range(2))
        r += f'<w:tr><w:trPr>{cs}</w:trPr>{cells}</w:tr>'
    return ('<w:tbl><w:tblPr><w:tblW w:w="9638" w:type="dxa"/><w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            f'</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4819"/><w:gridCol w:w="4819"/></w:tblGrid>{r}</w:tbl>')


def document(x, cant):
    body = ''.join(para(f'埋め草の行 {i + 1}') for i in range(34)) + (para('隙間', exact=x) if x > 0 else '') + table(cant_split=cant) + para('表のあとの段落')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{SECT}</w:body></w:document>'


mode = sys.argv[1] if len(sys.argv) > 1 else 'both'
import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for cant in ((True,) if mode == 'cant' else (False,) if mode == 'split' else (True, False)):
        for x in [int(v) for v in os.environ.get('ROWFIT_X', ','.join(str(i) for i in range(80, 101))).split(',')]:
            at = OUT / f'{"cant" if cant else "split"}{"" if CELL_SZ == 21 else "_sz%d" % CELL_SZ}_{x:02d}.docx'
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
                z.writestr('word/settings.xml', SETTINGS); z.writestr('word/document.xml', document(x, cant))
            d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
            try:
                t = d.Tables(1)
                c1 = t.Cell(1, 1).Range; l1 = d.Range(c1.Start, c1.Start)
                p2 = t.Cell(1, 1).Range.Paragraphs(2).Range; l2 = d.Range(p2.Start, p2.Start)
                c2 = t.Cell(2, 1).Range; r2 = d.Range(c2.Start, c2.Start)
                print(f'{"cant " if cant else "split"} X={x:2d} row1 line1 p{int(l1.Information(3))} y={l1.Information(6):.2f} line2 p{int(l2.Information(3))} y={l2.Information(6):.2f} | row2 p{int(r2.Information(3))} y={r2.Information(6):.2f}', flush=True)
            finally:
                d.Close(False)
finally:
    app.Quit()
