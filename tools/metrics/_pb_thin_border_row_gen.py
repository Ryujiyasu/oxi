# -*- coding: utf-8 -*-
"""How much does a THIN horizontal border (sz 1..7, eighths of a point) add to
an atLeast row's pitch?

en/reference__00f0e7f3 (compat 12, trHeight 400 atLeast, insideH sz 1): Word's
rows step 20.10pt (alternating 20.16 / 20.04 in the PDF), Oxi's 20.125 -- the
sz 1 border = 0.125pt -- so the last of 28 rows sits 0.5pt low.

Arms: insideH/top/bottom sz 1 / 2 / 3 / 4 / 5 / 6 / 7 x compat 12 / 15; 20 rows
of trHeight 400 atLeast, one short line each.
Readout: Word PDF baselines of the first and last row -> mean pitch.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/thin_border_row'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
N = 20


def settings(c):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{c}"/></w:compat></w:settings>')


def document(sz):
    b = ''.join(f'<w:{s} w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>'
                for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    ppr = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    rpr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="18"/></w:rPr>'
    rows = ''.join(f'<w:tr><w:trPr><w:trHeight w:val="400" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
                   f'<w:p>{ppr}<w:r>{rpr}<w:t>R{i:02d}</w:t></w:r></w:p></w:tc></w:tr>' for i in range(N))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblBorders>{b}</w:tblBorders></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>{rows}</w:tbl><w:p/>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for c, sz in itertools.product((12, 15), (1, 2, 3, 4, 5, 6, 7)):
            name = f'c{c}_sz{sz}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', settings(c))
                z.writestr('word/document.xml', document(sz))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            ys = sorted(round(l['spans'][0]['origin'][1], 3) for b in P.get_text('dict')['blocks'] for l in b.get('lines', [])
                        if l['spans'][0]['text'].startswith('R'))
            print(f'{name:10} first {ys[0]} last {ys[-1]} pitch {(ys[-1] - ys[0]) / (len(ys) - 1):.4f}', flush=True)
    finally:
        app.Quit()
