# -*- coding: utf-8 -*-
"""Floor sweep for _pb_floatlane: HOW WIDE must the lane beside a float be
before an EMPTY paragraph flows in it instead of below the float?

_pb_floatlane (Arial 11, leftFromText=142tw=7.1pt) bracketed it: lane 24pt ->
below, lane 30pt -> in the lane. This probe walks 1pt steps and varies BOTH
leftFromText (is the floor on the gross gap or on the gap net of the wrap
distance?) and the font size (is it a font-derived minimum?).

Usage:  python _pb_floatlane2_gen.py gen | measure
"""
import os, sys, zipfile, json

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _pb_floatlane_gen as P

OUTDIR = os.path.join(HERE, "..", "..", "pipeline_data", "_pb_floatlane2")

LEFT_MAR = 72.0
RIGHT_EDGE = 523.3
NROW = 8


def build(lane_pt, lft_tw, sz_hp):
    rpr = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="%d"/>' % sz_hp
    tblp_x = int(round((LEFT_MAR + lane_pt) * 20))
    width_tw = int(round((RIGHT_EDGE - (LEFT_MAR + lane_pt)) * 20))
    rows = ''.join(
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr>'
        % (width_tw, P.para('FROW%d' % i if i in (0, NROW - 1) else 'x'))
        for i in range(NROW))
    tbl = ('<w:tbl><w:tblPr>'
           '<w:tblpPr w:leftFromText="%d" w:rightFromText="%d" '
           'w:vertAnchor="text" w:horzAnchor="page" w:tblpX="%d" w:tblpY="1"/>'
           % (lft_tw, lft_tw, tblp_x) +
           '<w:tblOverlap w:val="never"/>'
           '<w:tblW w:w="%d" w:type="dxa"/><w:tblLayout w:type="fixed"/>' % width_tw +
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           '</w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="%d"/></w:tblGrid>%s</w:tbl>' % (width_tw, rows))
    body = (P.para('HEAD') + tbl + P.para('', rpr)
            + P.para('MARKER the quick brown fox jumps over the lazy dog again', rpr)
            + P.para('TAIL', rpr)
            + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
              'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document ' + P.W_NS + '><w:body>' + body + '</w:body></w:document>')


CASES = ([(l, 142, 22) for l in (18, 21, 24, 25, 26, 27, 28, 29, 30, 33, 36)]
         + [(l, 0, 22) for l in (12, 15, 18, 19, 20, 21, 22, 24, 27, 30)]
         + [(l, 284, 22) for l in (24, 27, 30, 31, 32, 33, 34, 36, 39)]
         + [(l, 142, 44) for l in (24, 27, 30, 33, 36, 42, 48)])


def name(l, lft, sz):
    return "f2_%03d_%03d_%02d" % (lft, int(round(l)), sz)


def gen():
    os.makedirs(OUTDIR, exist_ok=True)
    for l, lft, sz in CASES:
        with zipfile.ZipFile(os.path.join(OUTDIR, name(l, lft, sz) + '.docx'),
                             'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', P.CT)
            z.writestr('_rels/.rels', P.RELS)
            z.writestr('word/_rels/document.xml.rels', P.DOCRELS)
            z.writestr('word/document.xml', build(l, lft, sz))
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
                ftop = [r for r in rows if r['text'] == 'FROW0'][0]
                fbot = [r for r in rows if r['text'] == 'FROW7'][0]
                empt = [r for r in rows if r['text'] == '' and r['x'] < 200]
                mark = [r for r in rows if r['text'].startswith('MARKER')][0]
                lane = 'LANE' if empt and empt[0]['y'] < fbot['y'] else 'BELOW'
                mlane = 'LANE' if mark['y'] < fbot['y'] else 'BELOW'
                print("  %-14s ftop=%7.2f fbot=%7.2f empty=%s %-5s marker=%7.2f %-5s"
                      % (fn[:-5], ftop['y'], fbot['y'],
                         ("%7.2f" % empt[0]['y']) if empt else '   -   ', lane,
                         mark['y'], mlane))
            finally:
                d.Close(False)
    finally:
        word.Quit()
    json.dump(res, open(os.path.join(OUTDIR, '_result.json'), 'w'), indent=1)
    print('wrote _result.json')


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'gen'
    (gen if mode == 'gen' else measure)()
