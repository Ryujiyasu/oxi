# -*- coding: utf-8 -*-
"""How far below a CELL-bordered table does Word put the following paragraph?

technical__5175ec20 p8 (no-type docGrid linePitch=360, ASCII Times New Roman 12 /
eastAsia theme, tables drawn with tcBorders sz=4): comparing PDF rule lines against Oxi's
dump borders, every rule INSIDE a table agrees within 0.02, but each time the flow leaves
a table and re-enters one, Oxi loses height: bottom-rule 329.09 -> next top-rule 491.11 is
162.02 in Word and 161.10 in Oxi (-0.92), and 567.91 -> 597.10 is 29.19 against 28.50
(-0.69). Two such exits are the ~1.6pt that lets Oxi fit the 3.18 heading on page 8 where
Word pushes it to page 9. S1452 already adds the last row's bottom tcBorder width (0.5pt
at sz=4) when the next block is not a table, so either it is not firing here or the true
addend is larger.

Sheet: table A (one exact row, tcBorders only), then K text paragraphs, then table B.
Readout: the PDF rule ys, so A's bottom rule to B's top rule is measured the same way on
both sides, and the paragraph count separates a per-exit constant from a per-line error.

Arms: border weight (sz 4 / 12 / 24) x paragraphs between (1 / 3) x docGrid (no-type 360 /
absent) x the table's border source (tcBorders / tblBorders), the last arm being the
control that S1452 deliberately excludes.
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/tblexit'); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="ＭＳ 明朝" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
          '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style></w:styles>')


def bd(sz, tag):
    return f'<w:{tag} w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>'


def table(sz, label, source):
    if source == 'tbl':
        pr = ('<w:tblBorders>' + ''.join(bd(sz, t) for t in ('top', 'left', 'bottom', 'right')) + '</w:tblBorders>')
        tcb = ''
    else:
        pr = ''
        tcb = ('<w:tcBorders>' + ''.join(bd(sz, t) for t in ('top', 'left', 'bottom', 'right')) + '</w:tcBorders>')
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' + pr +
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
            '<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="306"/></w:trPr>'
            f'<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>{tcb}</w:tcPr>'
            f'<w:p><w:r><w:t>{label}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')


def document(sz, nparas, grid, source):
    mid = ''.join(f'<w:p><w:r><w:t>あいだの段落{i}</w:t></w:r></w:p>' for i in range(nparas))
    g = f'<w:docGrid w:linePitch="{grid}"/>' if grid else ''
    body = table(sz, 'A', source) + mid + table(sz, 'B', source) + '<w:p/>'
    sect = ('<w:sectPr><w:pgSz w:w="11907" w:h="16840" w:code="9"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720"/>'
            f'<w:cols w:space="425"/>{g}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


def rules(pdf):
    import pymupdf
    doc = pymupdf.open(pdf); p = doc[0]; ys = set()
    for dr in p.get_drawings():
        for it in dr['items']:
            if it[0] == 're' and it[1].height <= 2.0 and it[1].width > 50:
                ys.add(round(it[1].y0, 2))
            elif it[0] == 'l' and abs(it[1].y - it[2].y) < 0.5 and abs(it[1].x - it[2].x) > 50:
                ys.add(round(it[1].y, 2))
    doc.close()
    return sorted(ys)


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    arms = []
    for source in ('tc', 'tbl'):
        for sz in (4, 12, 24):
            arms.append((sz, 1, '360', source))
    arms += [(4, 3, '360', 'tc'), (4, 1, None, 'tc'), (4, 3, None, 'tc')]
    for sz, nparas, grid, source in arms:
        tag = f'{source}_sz{sz}_n{nparas}_g{grid}'
        at = OUT / f'{tag}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', STYLES)
            z.writestr('word/document.xml', document(sz, nparas, grid, source))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        pdf = str((OUT / f'{tag}.pdf').resolve())
        try:
            d.ExportAsFixedFormat(pdf, 17)
        finally:
            d.Close(False)
        r = rules(pdf)
        gap = round(r[2] - r[1], 2) if len(r) >= 3 else None
        print(f'{source:4} sz={sz:3} nparas={nparas} grid={str(grid):5} | rules={r[:4]} gap(A.bottom->B.top)={gap}', flush=True)
finally:
    app.Quit()
