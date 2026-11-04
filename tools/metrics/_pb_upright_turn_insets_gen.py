# -*- coding: utf-8 -*-
"""Do a quarter-turned upright text box's insets turn with the shape?

ja/creative__13152ea1: a roundRect with wp:extent 105.8 x 478, a:xfrm rot
5400000 and bodyPr upright="1" (default insets 7.2 / 3.6 / 7.2 / 3.6) sets its
horizontal lines 3.6pt left of and 3.7pt below where tIns-on-top / lIns-on-
the-left would put them.

Arms: a 100 x 300pt rect (unturned extent) with insets l 2 / t 10 / r 20 /
b 30pt, rot 0 / 90 / 270 (the 0 arm in its turned 300 x 100 extent), upright,
one left-aligned line «Inset» at Calibri 11. Readout from Word's PDF: the
box rect and the first glyph's x and baseline.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from _pb_shape_cmpd_shadow_gen import NS, CT, RR, DRELS, STYLES, EMU
OUT = 'tests/fixtures/upright_turn_insets'; os.makedirs(OUT, exist_ok=True)
X0, Y0 = 150.0, 150.0


def box(rot, ins=(2, 10, 20, 30), jc='left', anchor='t'):
    w, h = (300.0, 100.0) if rot == 0 else (100.0, 300.0)
    # place the TURNED box at (X0, Y0): an unturned extent's top-left sits (w-h)/2 off
    ox, oy = (X0, Y0) if rot == 0 else (X0 + (300 - 100) / 2, Y0 - (300 - 100) / 2)
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" '
            f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(ox * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(oy * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(w * EMU)}" cy="{int(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="2" name="B"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
            f'<wps:cNvSpPr/><wps:spPr><a:xfrm rot="{rot * 60000}"><a:off x="0" y="0"/><a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFCC"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>'
            f'<wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="{jc}"/></w:pPr><w:r><w:t>Inset</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
            f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="{int(ins[0] * EMU)}" tIns="{int(ins[1] * EMU)}" rIns="{int(ins[2] * EMU)}" bIns="{int(ins[3] * EMU)}" anchor="{anchor}" upright="1"/>'
            f'</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def document(rot, **kw):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body><w:p>{box(rot, **kw)}</w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        arms = [(r, {}) for r in (0, 90, 270)]
        if os.environ.get('PB_SET') == '2':
            arms = []
            for r in (90, 270):
                for k, ins in (('zero', (0, 0, 0, 0)), ('l20', (20, 0, 0, 0)), ('t20', (0, 20, 0, 0)), ('r20', (0, 0, 20, 0)), ('b20', (0, 0, 0, 20))):
                    for jc, an in (('left', 't'), ('right', 'b')):
                        arms.append((r, dict(ins=ins, jc=jc, anchor=an, tag=f'{k}_{jc}')))
        for rot, kw in arms:
            tag = kw.pop('tag', '')
            at = os.path.abspath(os.path.join(OUT, f'rot{rot}{"_" + tag if tag else ""}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/styles.xml', STYLES)
                z.writestr('word/document.xml', document(rot, **kw))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = fitz.open(pdf)
            rects = [[round(v, 2) for v in dr['rect']] for dr in P[0].get_drawings() if dr['rect'].width > 50]
            sp = [(round(s['bbox'][0], 2), round(s['bbox'][2], 2), round(s['origin'][1], 2)) for b in P[0].get_text('dict')['blocks']
                  for l in b.get('lines', []) for s in l['spans'] if s['text'].startswith('Inset')]
            print(f'rot{rot:<4}{tag:12} box {rects} text x/baseline {sp}', flush=True)
    finally:
        app.Quit()
