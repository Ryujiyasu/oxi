# -*- coding: utf-8 -*-
"""Does the FOLLOWING paragraph's left INDENT decide whether it flows in the
lane beside a float, or below it?

_pb_floatlane2 pinned the lane floor (net of leftFromText) at ~18.5pt and
showed that above it BOTH the empty paragraph and the long MARKER text flow in
the lane. ed025cbecffb's page-6 float leaves a 28pt net lane, its empty
paragraph flows there, but its following text (（記載上の注意）, left indent
31.5pt) goes BELOW. The indent is the only difference — this pins it.

  lane net 28.9pt, MARKER indent 0 / 9 / 21 / 31.5pt.

Usage:  python _pb_floatlane3_gen.py gen | measure
"""
import os, sys, zipfile, json

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_floatlane_gen as P

OUTDIR = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_floatlane3")

LEFT_MAR = 72.0
RIGHT_EDGE = 523.3
NROW = 8
LFT = 142


def build(lane_pt, ind_tw):
    rpr = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>'
    tblp_x = int(round((LEFT_MAR + lane_pt) * 20))
    width_tw = int(round((RIGHT_EDGE - (LEFT_MAR + lane_pt)) * 20))
    rows = ''.join(
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr>'
        % (width_tw, P.para('FROW%d' % i if i in (0, NROW - 1) else 'x'))
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
    ind = '<w:ind w:left="%d"/>' % ind_tw if ind_tw else ''
    marker = ('<w:p><w:pPr>' + P.SP0 + ind + '<w:rPr>' + rpr + '</w:rPr></w:pPr>'
              '<w:r><w:rPr>' + rpr + '</w:rPr>'
              '<w:t>MARKER the quick brown fox jumps over the lazy dog again</w:t>'
              '</w:r></w:p>')
    body = (P.para('HEAD') + tbl + P.para('', rpr) + marker + P.para('TAIL', rpr)
            + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
              'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document ' + P.W_NS + '><w:body>' + body + '</w:body></w:document>')


CASES = [(36, i) for i in (0, 180, 300, 420, 480, 540, 630)]


def name(l, i):
    return "f3_l%02d_i%04d" % (l, i)


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for l, i in CASES:
        with zipfile.ZipFile(os.path.join(OUTDIR, name(l, i) + '.docx'),
                             'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', P.CT)
            z.writestr('_rels/.rels', P.RELS)
            z.writestr('word/_rels/document.xml.rels', P.DOCRELS)
            z.writestr('word/document.xml', build(l, i))
            z.writestr('word/styles.xml', P.STYLES)
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
                                 'y': round(rs.Information(6), 2),
                                 'x': round(rs.Information(5), 2)})
                res[fn[:-5]] = rows
                fbot = [r for r in rows if r['text'] == 'FROW7'][0]
                empt = [r for r in rows if r['text'] == '' and r['x'] < 200][0]
                mark = [r for r in rows if r['text'].startswith('MARKER')][0]
                print("  %-14s fbot=%7.2f empty=%7.2f %-5s marker=%7.2f x=%6.2f %-5s"
                      % (fn[:-5], fbot['y'], empt['y'],
                         'LANE' if empt['y'] < fbot['y'] else 'BELOW',
                         mark['y'], mark['x'],
                         'LANE' if mark['y'] < fbot['y'] else 'BELOW'))
            finally:
                d.Close(False)
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, '_result.json'), 'w'), indent=1)
    print('wrote _result.json')


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    (gen if mode == 'gen' else measure)()
