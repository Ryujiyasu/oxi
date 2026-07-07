# -*- coding: utf-8 -*-
"""Adversarial probe: wrapSquare-anchored TEXTBOX (wps) side-wrap.

The S758 band machinery covers wrapSquare IMAGES; every corpus wrapSquare
anchor is a TEXTBOX (29dc6e x1, 2ea81a x4) which the registry ignores.
This probe pins the Word truth for the textbox variant: same shape as
probeximgfloat (right-aligned float at 第5条/第25条, text flows BESIDE),
but the anchor is a wps textbox with visible border + marker text.

Run: python tools/metrics/_probe_txbxwrap.py   (writes the docx)
Then: python tools/metrics/measure_pagination_word.py probeqtxbxwrap
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

esc = pg.esc
SENT = pg.SENT
MINCHO = pg.MINCHO
EMU = 12700

WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def rpr(sz="21"):
    return (f'<w:rFonts w:ascii="{MINCHO}" w:eastAsia="{MINCHO}" w:hAnsi="{MINCHO}"/>'
            f'<w:sz w:val="{sz}"/>')


def P(txt):
    r = rpr()
    return (f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{esc(txt)}</w:t></w:r></w:p>')


def anchor_txbx(pid, marker, cx_pt=142.0, cy_pt=113.0):
    cx, cy = int(cx_pt * EMU), int(cy_pt * EMU)
    r = rpr("18")
    inner = (f'<w:p><w:pPr><w:rPr>{r}</w:rPr></w:pPr>'
             f'<w:r><w:rPr>{r}</w:rPr><w:t>{esc(marker)}</w:t></w:r></w:p>')
    choice = (
        f'<w:drawing><wp:anchor {WP_NS} distT="0" distB="0" distL="114300" distR="114300" '
        'simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapSquare wrapText="bothSides"/>'
        f'<wp:docPr id="{pid}" name="TB{pid}"/><wp:cNvGraphicFramePr/>'
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


def build():
    parts = []
    for i in range(1, 51):
        if i in (5, 25):
            r = rpr()
            parts.append(
                f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{r}</w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:noProof/></w:rPr>{anchor_txbx(i, f"箱{i}")}</w:r>'
                f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">第{i}条　{esc(SENT)}</w:t></w:r></w:p>')
        else:
            parts.append(P(f"第{i}条　{SENT}"))
    body = "".join(parts) + pg.sectpr()
    out = pg.out("probeqtxbxwrap_wrapsquaretb.docx")
    pg.write_docx(out, pg.doc(body))
    print("wrote", out)


if __name__ == "__main__":
    build()
