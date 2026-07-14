# -*- coding: utf-8 -*-
"""Controlled sweep: wrapSquare textbox float with NEGATIVE posV
(positionV relativeFrom=paragraph, posOffset < 0) — the hmrc NI-strip
class that s758_tbs excludes via `tp.y >= 0.0`.

Derives:
  1. WHERE Word places the float vertically (float_top vs
     anchor_natural + posV; does prior-content overlap clamp it?)
  2. Where the ANCHOR paragraph lands (wrap below its own float?)
  3. Whether PRIOR paragraphs (above the anchor) are displaced
  4. Where FOLLOWERS land (wrap-below gap)

Structure per config page (mirrors the hmrc anchor XML: distT=0 distB=0
distL/R=114300, wrapSquare bothSides, wps rect with border):
  [exact filler][P0][P1][P2 exact-14 priors][optional exact gap]
  [ANCHOR para with float + text][F0][F1][F2]

cfg0 = control (no float) → natural chain positions.

Run: python tools/metrics/_pb_negfloat_gen.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.environ.get("TEMP", "."), 'negfloat')
os.makedirs(OUTDIR, exist_ok=True)
DOCX = os.path.join(OUTDIR, 'negfloat.docx')
PDF = os.path.join(OUTDIR, 'negfloat.pdf')
esc = pg.esc
EMU = 12700
FONT = "Calibri"

WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def rpr(sz="22"):
    return (f'<w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/>'
            f'<w:sz w:val="{sz}"/>')


def para(txt, spacing='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'):
    r = rpr()
    return (f'<w:p><w:pPr>{spacing}<w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')


def exact_para(txt, h_pt):
    return para(txt, f'<w:spacing w:after="0" w:line="{int(h_pt*20)}" w:lineRule="exact"/>')


def pagebreak():
    return ('<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:br w:type="page"/></w:r></w:p>')


def sq_txbx(pid, marker, pos_v_pt, cy_pt, cx_pt, pos_h_pt=20.0):
    cx, cy = int(cx_pt * EMU), int(cy_pt * EMU)
    posv, posh = int(pos_v_pt * EMU), int(pos_h_pt * EMU)
    r = rpr("14")
    inner = (f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>{r}</w:rPr></w:pPr>'
             f'<w:r><w:rPr>{r}</w:rPr><w:t>{esc(marker)}</w:t></w:r></w:p>')
    choice = (
        f'<w:drawing><wp:anchor {WP_NS} distT="0" distB="0" distL="114300" distR="114300" '
        'simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:posOffset>{posh}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{posv}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapSquare wrapText="bothSides"/>'
        f'<wp:docPr id="{pid}" name="NF{pid}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic {A_NS}>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
        '</wps:spPr>'
        f'<wps:txbx><w:txbxContent>{inner}</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="18000" tIns="18000" '
        'rIns="18000" bIns="18000" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>')
    return ('<mc:AlternateContent '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            f'<mc:Choice Requires="wps">{choice}</mc:Choice>'
            f'<mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>')


PAGE_TOP = 56.7    # 1134tw
PRIOR_H = 14.0     # exact prior lines
ANCHOR_NAT = 400.0 # target natural top of the anchor line

# (posV, cy, cx, gap_above)  — gap_above: exact spacer between priors and anchor
configs = [
    (None,   None, None, 0.0),    # 0 control
    (-10.0,  16.5, 400.0, 0.0),
    (-20.0,  16.5, 400.0, 0.0),
    (-33.44, 16.5, 400.0, 0.0),   # hmrc NI-strip value
    (-60.0,  16.5, 400.0, 0.0),
    (-100.0, 16.5, 400.0, 0.0),
    (-33.44, 40.0, 400.0, 0.0),   # box bottom overlaps the anchor line
    (-33.44, 60.0, 400.0, 0.0),   # box extends past anchor into followers
    (-33.44, 80.0, 400.0, 0.0),
    (-33.44, 16.5, 400.0, 60.0),  # float zone falls in EMPTY gap
    (-60.0,  16.5, 400.0, 60.0),
    (-33.44, 16.5, 216.4, 0.0),   # hmrc-faithful width (narrow, side room)
    (-33.44, 60.0, 216.4, 0.0),   # narrow + deep overlap (side-wrap possible)
]


def build():
    body = []
    pid = 100
    for idx, (pv, cy, cx, gap) in enumerate(configs):
        if idx:
            body.append(pagebreak())
        filler_h = ANCHOR_NAT - PAGE_TOP - 3 * PRIOR_H - gap
        body.append(exact_para(f'X{idx:02d}fill', filler_h))
        for j in range(3):
            body.append(exact_para(f'P{idx:02d}-{j}', PRIOR_H))
        if gap > 0:
            body.append(exact_para(f'G{idx:02d}gap', gap))
        r = rpr()
        if pv is None:
            body.append(para(f'A{idx:02d}anchor'))
        else:
            body.append(
                f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:noProof/></w:rPr>{sq_txbx(pid, f"B{idx:02d}box", pv, cy, cx)}</w:r>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">A{idx:02d}anchor</w:t></w:r></w:p>')
            pid += 1
        for j in range(3):
            body.append(para(f'F{idx:02d}-{j}'))
    body.append('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
                'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    pg.write_docx(DOCX, pg.doc(''.join(body)), font=FONT, sz="22")
    print('wrote', DOCX, len(configs), 'configs')


if __name__ == '__main__':
    build()
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
    doc.ExportAsFixedFormat(PDF, 17)
    doc.Close(False)
    word.Quit()
    import fitz
    d = fitz.open(PDF)
    sys.stdout.reconfigure(encoding='utf-8')
    print('pages:', len(d))
    # text lines
    locs = {}
    for p in range(len(d)):
        for b in d[p].get_text('dict')['blocks']:
            for l in b.get('lines', []):
                txt = ''.join(s['text'] for s in l.get('spans', [])).replace(' ', '')
                locs.setdefault(txt[:5], []).append((p + 1, round(l['bbox'][1], 2), round(l['bbox'][0], 2)))
    # box border rects per page (stroked rects ~ the float border)
    boxes = {}
    for p in range(len(d)):
        for dr in d[p].get_drawings():
            r0 = dr['rect']
            if r0.width > 50 and 5 < r0.height < 120:
                boxes.setdefault(p + 1, []).append((round(r0.y0, 2), round(r0.y1, 2), round(r0.x0, 2), round(r0.x1, 2)))
    print(f'{"cfg":>4} {"posV":>7} {"cy":>5} {"cx":>6} {"gap":>4} | box | P2(y,x) A(y,x) F0(y,x) F1(y,x) | dA dF0 dF1')
    ctrl = {}
    for idx, (pv, cy, cx, gap) in enumerate(configs):
        pgno = idx + 1
        def g(key):
            v = [e for e in locs.get(key[:5], []) if e[0] == pgno]
            return (v[0][1], v[0][2]) if v else (None, None)
        p2 = g(f'P{idx:02d}-2'); a = g(f'A{idx:02d}anchor'); f0 = g(f'F{idx:02d}-0'); f1 = g(f'F{idx:02d}-1')
        bx = boxes.get(pgno, [])
        bx_s = ' '.join(f'[{t}..{b} x{x0}]' for (t, b, x0, x1) in bx) or '-'
        if pv is None:
            ctrl = {'p2': p2[0], 'a': a[0], 'f0': f0[0], 'f1': f1[0]}
        def dd(v, k):
            if v[0] is None or ctrl.get(k) is None:
                return '     -'
            return f'{v[0] - ctrl[k]:+6.2f}'
        def yx(v):
            return f'({v[0]},{v[1]})'
        print(f'{idx:4d} {str(pv):>7} {str(cy):>5} {str(cx):>6} {gap:4.0f} | {bx_s} | '
              f'{yx(p2)} {yx(a)} {yx(f0)} {yx(f1)} | {dd(a, "a")} {dd(f0, "f0")} {dd(f1, "f1")}')
