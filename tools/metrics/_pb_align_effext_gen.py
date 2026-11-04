# -*- coding: utf-8 -*-
"""Does an ALIGNED floating shape (wp:align left/center/right, top/center/
bottom) align its box grown by wp:effectExtent, rather than the extent alone?

ja/reports__5373dd8a: the masthead roundRect (align left / top, margin;
effectExtent l=1.5 t=9 r=9.75 b=1.5pt for its outer shadow) sits at
(55.5, 81) in Word's PDF = margin + (l, t); the frame box (align left,
effectExtent l=3pt) at x 57 = margin + 3; its posOffset-placed twin (y
198.75) at exactly margin + 198.75. Oxi put both at the bare margin.

Arms: a 200 x 80pt rect text box, margin-relative, align h in left/center/
right x align v in top/center/bottom, with (a) effectExtent l=2 t=9 r=5 b=3pt
and no effect, (b) the same extent and an outerShdw, (c) the outerShdw and a
zero extent; plus a posOffset arm with the extent. Readout: the box's fill
rect in Word's PDF.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from _pb_shape_cmpd_shadow_gen import NS, CT, RR, DRELS, STYLES, EMU
OUT = 'tests/fixtures/align_effext'; os.makedirs(OUT, exist_ok=True)
W, H = 200.0, 80.0
EXT = (2.0, 9.0, 5.0, 3.0)


def box(h, v, ext, shadow):
    ph = (f'<wp:align>{h}</wp:align>' if h != 'off' else f'<wp:posOffset>{40 * EMU}</wp:posOffset>')
    pv = (f'<wp:align>{v}</wp:align>' if v != 'off' else f'<wp:posOffset>{40 * EMU}</wp:posOffset>')
    l, t, r, b = ext
    eff = (f'<a:effectLst><a:outerShdw dist="{9 * EMU}" dir="2700000" algn="bl" rotWithShape="0">'
           f'<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>') if shadow else ''
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" '
            f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="margin">{ph}</wp:positionH>'
            f'<wp:positionV relativeFrom="margin">{pv}</wp:positionV>'
            f'<wp:extent cx="{int(W * EMU)}" cy="{int(H * EMU)}"/>'
            f'<wp:effectExtent l="{int(l * EMU)}" t="{int(t * EMU)}" r="{int(r * EMU)}" b="{int(b * EMU)}"/><wp:wrapNone/>'
            f'<wp:docPr id="2" name="B"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
            f'<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(W * EMU)}" cy="{int(H * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="3366CC"/></a:solidFill><a:ln><a:noFill/></a:ln>{eff}</wps:spPr>'
            f'<wps:txbx><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
            f'<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def document(arm):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body><w:p>{box(*arm)}</w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


ARMS = []
for kind, ext, sh in (('ext', EXT, False), ('extsh', EXT, True), ('sh', (0, 0, 0, 0), True)):
    for h in ('left', 'center', 'right'):
        for v in ('top', 'center', 'bottom'):
            ARMS.append((kind, (h, v, ext, sh)))
ARMS.append(('ext', ('off', 'off', EXT, False)))

if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for kind, arm in ARMS:
            name = f'{kind}_{arm[0]}_{arm[1]}'
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/styles.xml', STYLES)
                z.writestr('word/document.xml', document(arm))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = fitz.open(pdf)
            rects = [[round(v, 2) for v in dr['rect']] for dr in P[0].get_drawings()
                     if dr.get('fill') and dr['fill'][2] > 0.7 and dr['fill'][0] < 0.3]
            print(f'{name:22} {rects}', flush=True)
    finally:
        app.Quit()
