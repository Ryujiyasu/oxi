# -*- coding: utf-8 -*-
"""Which width does a FIXED-layout table take when tblW disagrees with its grid?

ja/policies__07543a6b p46: tblLayout fixed, tblW 5043 pct (419.22pt of a 415.65
text width), tblGrid 2665/754/2546/2399 = 8364tw (418.2pt), cells tcW in pct.
Word's rules sit on the grid (col 1 133.23 = 2665tw); Oxi scales the grid up to
tblW (133.58) and a cell line keeps a character Word wraps.

Arms: layout fixed / autofit x tblW pct 5043 / pct 4500 / dxa 9000 / dxa 7800 /
auto, grid 2665/754/2546/2399, tcW pct (the document's) or dxa (= the grid).
Readout: Word's vertical rule centres from the PDF.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/fixed_tblw_grid'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SET = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
       '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
GRID = [2665, 754, 2546, 2399]
PCT = [1593, 451, 1522, 1434]
TBLW = {'pct5043': ('5043', 'pct'), 'pct4500': ('4500', 'pct'), 'dxa9000': ('9000', 'dxa'), 'dxa7800': ('7800', 'dxa'), 'auto': ('0', 'auto')}


def document(layout, tblw, tcw, sz=8, ind=-10):
    w, t = TBLW[tblw]
    b = ''.join(f'<w:{s} w:val="{'single' if sz else 'nil'}" w:sz="{sz}" w:space="0" w:color="000000"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    lay = '<w:tblLayout w:type="fixed"/>' if layout == 'fixed' else ''
    cells = ''.join(
        f'<w:tc><w:tcPr><w:tcW w:w="{PCT[i] if tcw == "pct" else GRID[i]}" w:type="{tcw}"/></w:tcPr><w:p><w:r><w:t>c{i}</w:t></w:r></w:p></w:tc>'
        for i in range(4))
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="{w}" w:type="{t}"/><w:tblInd w:w="{ind}" w:type="dxa"/><w:tblBorders>{b}</w:tblBorders>{lay}'
           '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{g}"/>' for g in GRID) + f'</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p/>{tbl}<w:p/>'
            '<w:sectPr><w:pgSz w:w="11907" w:h="16840"/><w:pgMar w:top="1418" w:right="1797" w:bottom="1276" w:left="1797" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        print('text 89.85 .. 505.5 (415.65); grid 133.25 / 37.70 / 127.30 / 119.95 = 418.2')
        arms = [('fixed', tw, 'dxa', sz, ind) for tw in ('pct5043', 'pct4500') for sz in (0, 8, 24) for ind in (0, -10, -200, 200)]
        for layout, tblw, tcw, sz, ind in arms:
                    name = f'{layout}_{tblw}_{tcw}_sz{sz}_ind{ind}'
                    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
                    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                        z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', SET)
                        z.writestr('word/document.xml', document(layout, tblw, tcw, sz, ind))
                    pdf = at[:-5] + '.pdf'
                    d = app.Documents.Open(at, ReadOnly=True)
                    try:
                        d.ExportAsFixedFormat(pdf, 17)
                    finally:
                        d.Close(False)
                    P = pymupdf.open(pdf)[0]
                    xs = sorted(set(round((g['rect'].x0 + g['rect'].x1) / 2, 2) for g in P.get_drawings() if g['rect'].height > 5 and g['rect'].width < 3))
                    if not xs:
                        dd = app.Documents.Open(at, ReadOnly=True)
                        try:
                            t = dd.Tables(1)
                            xs = [round(t.Rows(1).Cells(1).Range.Information(5), 2), round(t.Rows(1).Cells(4).Range.Information(5), 2)]
                            cw = [round(t.Rows(1).Cells(i).Width, 2) for i in range(1, 5)]
                            print(f'{name:30} COM cell x {xs} widths {cw} total {round(sum(cw), 2)}', flush=True); continue
                        finally:
                            dd.Close(False)
                    widths = [round(b - a, 2) for a, b in zip(xs, xs[1:])]
                    print(f'{name:30} rules {xs}  widths {widths}  total {round(xs[-1] - xs[0], 2) if xs else None}', flush=True)
    finally:
        app.Quit()
