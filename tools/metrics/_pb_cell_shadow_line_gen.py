# -*- coding: utf-8 -*-
"""Does a w14:shadow run grow its line inside a TABLE CELL the way it does in
the body (OXI_TEXT_SHADOW_LINE: extents above max(0, blur - dy), below
max(0, blur + dy), times 0.05 x fs^0.70)?

en/policies__009625de: the header table's «2016» (Cambria bold 18, w14:shadow
blur 4 / dist 3 / dir 90) -- Word's body starts 3.1pt below Oxi's, the
predicted growth (1 + 7) x 0.05 x 18^0.7 = 3.02.

Arms: a one-row bordered table (cell margins 0) holding one Calibri / Cambria
line at 12 / 18 / 24pt, the run with no shadow or w14:shadow blur 4 dist 3 at
dir 90 (down) / 270 (up), then a body paragraph. Readout from Word's PDF: the
table's bottom border y and the run's baseline.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/cell_shadow_line'; os.makedirs(OUT, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')


def shadow(d):
    if d is None:
        return ''
    return (f'<w14:shadow w14:blurRad="50800" w14:dist="38100" w14:dir="{d * 60000}" w14:sx="100000" w14:sy="100000" '
            'w14:kx="0" w14:ky="0" w14:algn="tl"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:shadow>')


def document(font, size, d, rule='auto'):
    rpr = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{size * 2}"/>{shadow(d)}'
    sp = ('<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>' if rule == 'auto' else
          f'<w:spacing w:after="0" w:line="{rule[7:]}" w:lineRule="atLeast"/>')
    b = '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    tbl = (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblBorders>{b}</w:tblBorders>'
           '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
           '</w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>'
           f'<w:p><w:pPr>{sp}</w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>Shadow 2016</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{tbl}'
            f'<w:p><w:pPr>{sp}</w:pPr><w:r><w:t>After</w:t></w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


ARMS = [(f, s, d, 'auto') for f in ('Calibri', 'Cambria') for s in (12, 18, 24) for d in (None, 90, 270)]
if os.environ.get('PB_ATLEAST'):
    ARMS = [('Arial', 14, d, r) for r in ('atLeast0', 'atLeast300', 'atLeast400') for d in (None, 90)]

if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for f, s, d, rule in ARMS:
            name = f'{f}_{s}_{"none" if d is None else d}' + ('' if rule == 'auto' else '_' + rule)
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/document.xml', document(f, s, d, rule))
            pdf = at[:-5] + '.pdf'
            doc = app.Documents.Open(at, ReadOnly=True)
            try:
                doc.ExportAsFixedFormat(pdf, 17)
            finally:
                doc.Close(False)
            P = fitz.open(pdf)
            hs = sorted(round(dr['rect'].y0, 2) for dr in P[0].get_drawings() if dr['rect'].width > 200)
            ys = {s2['text'][:6]: round(s2['origin'][1], 2) for b in P[0].get_text('dict')['blocks']
                  for l in b.get('lines', []) for s2 in l['spans']}
            print(f'{name:18} rules {hs} base {ys.get("Shadow")} after {ys.get("After")}', flush=True)
    finally:
        app.Quit()
