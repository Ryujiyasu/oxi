# -*- coding: utf-8 -*-
"""Do EMPTY paragraphs after a wrap-below floating table flow in the side LANE,
or below the float? And how wide must the lane be?

ed025cbecffb (Word COM + PNG): its page-6 float (vertAnchor=text,
horzAnchor=page, tblpX=2008tw -> left lane 35.1pt) is followed by ONE empty
paragraph; Word puts it at y=78 (beside the float, x=65.25) and the next TEXT
paragraph at 289.5 = the float's bottom. Oxi's wrap-below branch spends a line
below the float on that empty -> everything after is 18pt low.

This probe sweeps the LANE width with a fixed float and N empty paragraphs
between the float and a long MARKER line.

  empty y ~ float top      -> the empty flows in the lane
  empty y ~ float bottom   -> the empty is pushed below

Usage:  python _pb_floatlane_gen.py gen | measure
"""
import os, sys, zipfile, json

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_pb_floatlane")

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
          '<w:styles ' + W_NS + '>'
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
LEFT_MAR = 72.0          # 1440tw
RIGHT_EDGE = 523.3       # page 595.3 - 72
NROW = 8                 # float rows (~8 lines tall)
LFT = 142                # leftFromText, as ed025 states it


def para(text="", rpr=R11):
    if text:
        return ('<w:p><w:pPr>' + SP0 + '<w:rPr>' + rpr + '</w:rPr></w:pPr>'
                '<w:r><w:rPr>' + rpr + '</w:rPr><w:t>' + text + '</w:t></w:r></w:p>')
    return '<w:p><w:pPr>' + SP0 + '<w:rPr>' + rpr + '</w:rPr></w:pPr></w:p>'


def build(lane_pt, n_empty):
    """lane_pt = free width between the left margin and the float's left edge."""
    tblp_x = int(round((LEFT_MAR + lane_pt) * 20))
    width_tw = int(round((RIGHT_EDGE - (LEFT_MAR + lane_pt)) * 20))
    rows = ''.join(
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr>'
        % (width_tw, para('FROW%d' % i if i in (0, NROW - 1) else 'x'))
        for i in range(NROW))
    tbl = ('<w:tbl><w:tblPr>'
           '<w:tblpPr w:leftFromText="%d" w:rightFromText="%d" '
           'w:vertAnchor="text" w:horzAnchor="page" w:tblpX="%d" w:tblpY="1"/>'
           % (LFT, LFT, tblp_x) +
           '<w:tblOverlap w:val="never"/>'
           '<w:tblW w:w="%d" w:type="dxa"/><w:tblLayout w:type="fixed"/>' % width_tw +
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>%s</w:tbl>' % (width_tw, rows))
    body = (para('HEAD') + tbl
            + ''.join(para() for _ in range(n_empty))
            + para('MARKER the quick brown fox jumps over the lazy dog again')
            + para('TAIL')
            + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
              'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document ' + W_NS + '><w:body>' + body + '</w:body></w:document>')


LANES = [0, 6, 12, 18, 24, 30, 36, 42, 48, 60]
NEMPTY = [1, 3]
CASES = [(l, n) for n in NEMPTY for l in LANES]


def name(l, n):
    return "fl_n%d_l%02d" % (n, l)


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for l, n in CASES:
        with zipfile.ZipFile(os.path.join(OUTDIR, name(l, n) + '.docx'),
                             'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/document.xml', build(l, n))
            z.writestr('word/styles.xml', STYLES)
    print('generated %d -> %s' % (len(CASES), os.path.abspath(OUTDIR)))


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
                rows = []
                for i in range(1, d.Paragraphs.Count + 1):
                    r = d.Paragraphs(i).Range
                    txt = r.Text.replace('\r', '').replace('\x07', '')
                    rs = d.Range(r.Start, r.Start)
                    rows.append({'i': i, 'text': txt[:18],
                                 'page': rs.Information(3),
                                 'y': round(rs.Information(6), 2),
                                 'x': round(rs.Information(5), 2)})
                res[fn[:-5]] = rows
                def g(pred):
                    hits = [r for r in rows if pred(r)]
                    return hits[0] if hits else None
                head = g(lambda r: r['text'].startswith('HEAD'))
                ftop = g(lambda r: r['text'] == 'FROW0')
                fbot = g(lambda r: r['text'] == 'FROW7')
                empt = [r for r in rows if r['text'] == '']
                mark = g(lambda r: r['text'].startswith('MARKER'))
                print("  %-14s HEAD=%s FROW0=%s FROW7=%s empties=%s MARKER=%s" % (
                    fn[:-5],
                    head['y'] if head else '-', ftop['y'] if ftop else '-',
                    fbot['y'] if fbot else '-',
                    [(e['y'], e['x']) for e in empt],
                    (mark['y'], mark['x']) if mark else '-'))
            finally:
                d.Close(False)
    finally:
        word.Quit()
    out = os.path.join(OUTDIR, '_result.json')
    json.dump(res, open(out, 'w'), indent=1)
    print('wrote', out)


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    (gen if mode == 'gen' else measure)()
