# -*- coding: utf-8 -*-
"""How far does a compound / art-like CELL top border push the cell text down?

ja/administrative__0ba21b0c: tcBorders top dashDotStroked sz4 -- Word's band
178.5..180.4 and the first line top 181.5 (3.0 below the row top); Oxi took the
border as sz/8 = 0.5. The paragraph-border law (OXI_PBDR_COMPOUND) has
dashDotStroked = 3 fixed, but cell borders were never probed.

Each document: a 12pt exact paragraph «前», then a one-cell table whose cell's
top border is the arm, with an exact-12pt paragraph «中»; readout = COM
Information(6) of «中» minus that of «前» (= 12 + the border's push).
Arms: style x sz, placed on tcBorders and on tblBorders.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/cell_compound_border'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
PPR = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>'
STYLES = ['single', 'dashDotStroked', 'double', 'thinThickSmallGap', 'threeDEmboss']
SIZES = [4, 12, 24]


def p(t):
    return f'<w:p>{PPR}<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/></w:rPr><w:t>{t}</w:t></w:r></w:p>'


def document(style, sz, level):
    b = f'<w:top w:val="{style}" w:sz="{sz}" w:space="0" w:color="000000"/>'
    tbl_b = f'<w:tblBorders>{b}</w:tblBorders>' if level == 'tbl' else ''
    tc_b = f'<w:tcBorders>{b}</w:tcBorders>' if level == 'tc' else ''
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/>{tbl_b}</w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>'
           f'<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/>{tc_b}</w:tcPr>{p("中")}</w:tc></w:tr></w:tbl>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{p("前")}{tbl}{p("後")}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for style, sz, level in itertools.product(STYLES, SIZES, ('tc', 'tbl')):
            name = f'{style}_{sz}_{level}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/document.xml', document(style, sz, level))
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                ys = []
                for i in range(1, d.Paragraphs.Count + 1):
                    r = d.Paragraphs(i).Range
                    ys.append((r.Text.strip(chr(13) + chr(7) + " "), d.Range(r.Start, r.Start).Information(6)))
                y = dict((t, v) for t, v in ys if t)
            finally:
                d.Close(False)
            print(f'{name:32} push {y["中"] - y["前"] - 12:6.2f}  after {y["後"] - y["中"] - 12:6.2f}  (sz/8 {sz / 8:.2f})', flush=True)
    finally:
        app.Quit()
