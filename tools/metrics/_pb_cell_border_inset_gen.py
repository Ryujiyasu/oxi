# -*- coding: utf-8 -*-
"""Does a cell's LEFT border width push its text right, and does a thick TOP
border grow the first row?

ja/forms__030ed901 (compat 15, tblInd 108, no tblCellMar): first-row cells
override the left / top border to sz 12 (1.5pt). Word's cell text starts 0.8pt
right of Oxi's in every row, and Word's first row is 1.3pt shorter.

Arms: left border sz 4 / 12 / 24 on the first cell (tcBorders) or the table
(tblBorders) x top border sz 4 / 12 / 24 x compat 14 / 15.
Readout: Word PDF text x of each cell, border rectangles, row boundaries.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/cell_border_inset'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')


def settings(c):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{c}"/></w:compat></w:settings>')


def document(left, top, where):
    b = lambda side, sz: f'<w:{side} w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>'
    tbl_b = ''.join(b(s, 4) for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    if where == 'tbl':
        tbl_b = b('top', top) + b('left', left) + ''.join(b(s, 4) for s in ('bottom', 'right', 'insideH', 'insideV'))
        tc_b = ''
    else:
        tc_b = f'<w:tcBorders>{b("top", top)}{b("left", left)}</w:tcBorders>'
    ppr = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    rpr = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="18"/></w:rPr>'
    cell = lambda t, tcb: f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>{tcb}</w:tcPr><w:p>{ppr}<w:r>{rpr}<w:t>{t}</w:t></w:r></w:p></w:tc>'
    rows = (f'<w:tr>{cell("ROWAONE", tc_b)}{cell("ROWATWO", "")}</w:tr>'
            f'<w:tr>{cell("ROWBONE", "")}{cell("ROWBTWO", "")}</w:tr>'
            f'<w:tr>{cell("ROWCONE", "")}{cell("ROWCTWO", "")}</w:tr>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="108" w:type="dxa"/><w:tblBorders>{tbl_b}</w:tblBorders></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>{rows}</w:tbl><w:p/>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for c, where, left, top in itertools.product((15, 14), ('tc', 'tbl'), (4, 12, 24), (4, 12, 24)):
            if left != top and left != 4 and top != 4:
                continue
            name = f'c{c}_{where}_l{left}_t{top}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', settings(c))
                z.writestr('word/document.xml', document(left, top, where))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            txt = {}
            for b in P.get_text('rawdict')['blocks']:
                for l in b.get('lines', []):
                    t = ''.join(ch['c'] for s in l['spans'] for ch in s['chars'])
                    if t.startswith('ROW'):
                        txt[t[:7]] = (round(l['spans'][0]['chars'][0]['bbox'][0], 2), round(l['spans'][0]['origin'][1], 2))
            vx, hy = set(), set()
            for dr in P.get_drawings():
                for it in dr['items']:
                    if it[0] == 're':
                        r = it[1]
                        if r.width < r.height: vx.add((round(r.x0, 2), round(r.width, 2)))
                        else: hy.add((round(r.y0, 2), round(r.height, 2)))
            print(f'{name:16} A1{txt.get("ROWAONE")} A2{txt.get("ROWATWO")} B1{txt.get("ROWBONE")} C1{txt.get("ROWCONE")} '
                  f'vx{sorted(vx)[:4]} hy{sorted(hy)[:4]}', flush=True)
    finally:
        app.Quit()
