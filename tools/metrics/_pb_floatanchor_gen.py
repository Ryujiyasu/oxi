# -*- coding: utf-8 -*-
"""Which PAGE does a vertAnchor="page" floating table render on when its
anchor position nears the page bottom? (the S869/S870/S871 bundle blocker)

policies__0009e9db: anchor room 1.7pt  -> Word floats on the NEXT page.
459f05 (canary):    anchor room 19.0pt -> Word floats on the SAME page.
Hypothesis: the float anchors on the page where the NEXT LINE would fit
(threshold = the following content's line height). This probe pins the flip
point and its font-size dependence.

METHOD (1 render per config): body = 53 Arial-11 filler lines + ONE empty
spacer para with line=X lineRule=exact (anchor_y = 72 + 53*12.6489 + X/20,
room = 769.9 - anchor_y = 27.55 - X/20) + a floating table (vertAnchor=page
tblpY=2233 = 111.65pt, one short row) + a MARKER para.
Series m11: MARKER Arial 11 (line 12.649). Series m22: MARKER Arial 22
(line 25.30). If the flip tracks the marker size -> threshold = following
para's line height; if not -> a fixed minimum.
Read per config: the page of FLOATCELL and of MARKER (Word COM Info(3)/(6)).

Usage:
  python _pb_floatanchor_gen.py gen
  python _pb_floatanchor_gen.py measure
"""
import os, sys, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_floatanchor")

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:styles {W_NS}>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/>'
          '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>'
          '</w:style></w:styles>')

R11 = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
R22 = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="44"/>'
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
ADV = 12.6489       # Arial 11 hhea line (S805)
NFILL = 53
TOP = 72.0
CBOT = 769.9        # A4 16838tw, bottom 1440tw (the policies geometry)


def build(x_tw, marker_sz22):
    fill = ''.join(
        f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{R11}</w:rPr><w:t>L{i:02d} alpha beta gamma.</w:t></w:r></w:p>'
        for i in range(NFILL))
    spacer = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" '
              f'w:line="{x_tw}" w:lineRule="exact"/><w:rPr>{R11}</w:rPr></w:pPr></w:p>') if x_tw else ''
    tbl = ('<w:tbl><w:tblPr>'
           '<w:tblpPr w:leftFromText="180" w:rightFromText="180" '
           'w:vertAnchor="page" w:horzAnchor="margin" w:tblpY="2233"/>'
           '<w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
           '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
           f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
           f'<w:r><w:rPr>{R11}</w:rPr><w:t>FLOATCELL</w:t></w:r></w:p>'
           '</w:tc></w:tr></w:tbl>')
    rm = R22 if marker_sz22 else R11
    marker = (f'<w:p><w:pPr>{SP0}<w:rPr>{rm}</w:rPr></w:pPr>'
              f'<w:r><w:rPr>{rm}</w:rPr><w:t>MARKER</w:t></w:r></w:p>')
    body = (fill + spacer + tbl + marker +
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
            'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


def name(x, m22):
    return f"fa_{'m22' if m22 else 'm11'}_{x:04d}"


CASES = [(x, m) for m in (False, True)
         for x in (0, 100, 200, 240, 280, 320, 360, 400, 440, 480, 520)]


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for x, m in CASES:
        with zipfile.ZipFile(os.path.join(OUTDIR, name(x, m) + '.docx'),
                             'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', build(x, m))
            z.writestr('word/styles.xml', STYLES)
    print(f'generated {len(CASES)} -> {os.path.abspath(OUTDIR)}')


def measure():
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for fn in sorted(f for f in os.listdir(OUTDIR) if f.endswith('.docx')):
            p = os.path.abspath(os.path.join(OUTDIR, fn))
            d = word.Documents.Open(p, ReadOnly=True)
            try:
                fc = mk = None
                for i in range(1, d.Paragraphs.Count + 1):
                    r = d.Paragraphs(i).Range
                    t = ''.join(ch for ch in r.Text if ch.isalnum())
                    if t in ('FLOATCELL', 'MARKER'):
                        rs = d.Range(r.Start, r.Start)
                        pg = rs.Information(3)   # wdActiveEndPageNumber via collapsed start
                        y = rs.Information(6)
                        if t == 'FLOATCELL':
                            fc = (pg, round(y, 2))
                        else:
                            mk = (pg, round(y, 2))
                res[fn[:-5]] = {'float': fc, 'marker': mk}
                print(f"  {fn[:-5]}: float {fc}   marker {mk}   (paras={d.Paragraphs.Count})")
            finally:
                d.Close(False)
    finally:
        word.Quit()
    out = os.path.join(OUTDIR, '_result.json')
    json.dump(res, open(out, 'w'), indent=1)
    print('wrote', out)
    # readout
    print('\nx_tw  room(pt)  m11 float-page   m22 float-page')
    for x in sorted({c[0] for c in CASES}):
        room = CBOT - (TOP + NFILL * ADV + x / 20.0)
        a = res.get(f'fa_m11_{x:04d}', {}).get('float')
        b = res.get(f'fa_m22_{x:04d}', {}).get('float')
        print(f'{x:5d} {room:8.2f}   {a}   {b}')


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen':
        gen()
    elif mode == 'measure':
        measure()
