# -*- coding: utf-8 -*-
"""Derive Word's placement of a floating table with w:tblpYSpec="bottom"
(policies__003496577: Table 3-1, horzAnchor=margin, NO vertAnchor, tblpYSpec
=bottom; the table renders ~60pt LOWER than Oxi's flow placement, tipping one
row across the p10/p11 split). Oxi does not parse tblpYSpec at all.

METHOD (1 render per config): K Arial-11 filler lines + an anchor paragraph +
a floating table (tblpYSpec=bottom, R rows of exact height H) + a MARKER para.
Read the table's first-cell page/y and the marker page/y (Word COM Info(3)/(6),
collapsed start). Geometry: A4, top margin 72pt, content bottom 769.9pt,
Arial-11 line 12.6489pt.

Series:
  fit_none  : small table (3x20pt=60pt), NO vertAnchor      -> is row1 at CBOT-60 (page-bottom)?
  fit_text  : small table, vertAnchor="text"                -> bottom of text region?
  fit_marg  : small table, vertAnchor="margin"
  big_none  : policies 7 rows (540pt), NO vertAnchor        -> where does row1 start + split?
Vary K (filler) to see if the position is anchor-independent (pure page bottom)
or anchor-relative.

Usage:  python _pb_tblpybottom_gen.py gen   |   measure
"""
import os, sys, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_tblpybottom")
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
SP0 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
ADV = 12.6489
TOP = 72.0
CBOT = 769.9   # A4 16838tw - 1440tw bottom margin

# policies Table 3-1 exact row heights (twip)
BIG_ROWS = [1134, 1862, 1386, 1675, 1249, 2099, 1400]


def tbl(rows_tw, v_anchor):
    va = f'w:vertAnchor="{v_anchor}" ' if v_anchor else ''
    trs = ''
    for k, h in enumerate(rows_tw):
        trs += (f'<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="{h}"/></w:trPr>'
                '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
                f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{R11}</w:rPr><w:t>ROW{k}</w:t></w:r></w:p></w:tc></w:tr>')
    return ('<w:tbl><w:tblPr>'
            f'<w:tblpPr w:horzAnchor="margin" {va}w:tblpYSpec="bottom"/>'
            '<w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
            '</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
            + trs + '</w:tbl>')


def build(k_fill, rows_tw, v_anchor):
    fill = ''.join(
        f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{R11}</w:rPr><w:t>L{i:02d} alpha beta gamma.</w:t></w:r></w:p>'
        for i in range(k_fill))
    anchor = (f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
              f'<w:r><w:rPr>{R11}</w:rPr><w:t>ANCHORLINE</w:t></w:r></w:p>')
    marker = (f'<w:p><w:pPr>{SP0}<w:rPr>{R11}</w:rPr></w:pPr>'
              f'<w:r><w:rPr>{R11}</w:rPr><w:t>MARKER</w:t></w:r></w:p>')
    body = (fill + anchor + tbl(rows_tw, v_anchor) + marker +
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
            'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {W_NS}><w:body>{body}</w:body></w:document>')


SMALL = [400, 400, 400]  # 3 x 20pt = 60pt
CASES = []
for k in (5, 20, 40):
    CASES.append((f"fit_none_k{k:02d}", k, SMALL, None))
    CASES.append((f"fit_text_k{k:02d}", k, SMALL, "text"))
    CASES.append((f"fit_marg_k{k:02d}", k, SMALL, "margin"))
    CASES.append((f"big_none_k{k:02d}", k, BIG_ROWS, None))


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for nm, k, rows, va in CASES:
        with zipfile.ZipFile(os.path.join(OUTDIR, nm + '.docx'), 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', build(k, rows, va))
            z.writestr('word/styles.xml', STYLES)
    print(f'generated {len(CASES)} -> {os.path.abspath(OUTDIR)}')


def measure():
    import win32com.client
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    res = {}
    try:
        for nm, k, rows, va in CASES:
            p = os.path.abspath(os.path.join(OUTDIR, nm + '.docx'))
            d = word.Documents.Open(p, ReadOnly=True)
            try:
                anchor = row0 = mk = None
                # anchor + marker via paragraphs; table row0 via the first table
                for i in range(1, d.Paragraphs.Count + 1):
                    r = d.Paragraphs(i).Range
                    t = ''.join(ch for ch in r.Text if ch.isalnum())
                    if t in ('ANCHORLINE', 'MARKER'):
                        rs = d.Range(r.Start, r.Start)
                        v = (rs.Information(3), round(rs.Information(6), 1))
                        if t == 'ANCHORLINE': anchor = v
                        else: mk = v
                if d.Tables.Count >= 1:
                    c = d.Tables(1).Cell(1, 1).Range
                    cs = d.Range(c.Start, c.Start)
                    row0 = (cs.Information(3), round(cs.Information(6), 1))
                    # bottom of last row
                    lastr = d.Tables(1).Rows.Count
                    cl = d.Tables(1).Cell(lastr, 1).Range
                    cls = d.Range(cl.Start, cl.Start)
                    rowN = (cls.Information(3), round(cls.Information(6), 1))
                else:
                    rowN = None
                res[nm] = {'anchor': anchor, 'row0': row0, 'rowN': rowN, 'marker': mk}
                print(f"  {nm}: anchor {anchor}  row0 {row0}  rowN {rowN}  marker {mk}")
            finally:
                d.Close(False)
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, '_result.json'), 'w'), indent=1)
    print("wrote _result.json")


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    if mode == 'gen': gen()
    elif mode == 'measure': measure()
