# -*- coding: utf-8 -*-
"""縦中横 (eastAsianLayout w:vert="1") inside a vertical (eaVert) text box: how
wide is the horizontal run drawn, and how far down does it advance?

ja/reports__542222d0: «２４» (HGP創英角ﾎﾟｯﾌﾟ体 36 bold, w 88%, vert=1
vertCompress=1) at the head of an eaVert box column -- Word draws the two digits
side by side, ink ~46pt wide (the 46.7pt column, 83/64 x 36) and about one em
tall; Oxi stacked them.

Arms: an eaVert box (300 x 300pt), one paragraph «あ[TCY]いう», TCY in {２４, 12,
１２３} at ＭＳ ゴシック 20 / 36pt, vertCompress 0 / 1, w 100 / 88. Readout from
Word's PDF: the TCY glyphs' bboxes and the next character's top.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from _pb_shape_cmpd_shadow_gen import NS, CT, RR, DRELS, STYLES, EMU
OUT = 'tests/fixtures/tcy_vert'; os.makedirs(OUT, exist_ok=True)
X0, Y0, W, H = 150.0, 100.0, 300.0, 300.0


def run(t, fs, extra=''):
    return (f'<w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/>'
            f'<w:sz w:val="{fs * 2}"/>{extra}</w:rPr><w:t>{t}</w:t></w:r>')


def box(tcy, fs, comp, wpct):
    lay = f'<w:eastAsianLayout w:id="1" w:vert="1"{" w:vertCompress=\"1\"" if comp else ""}/>'
    wx = f'<w:w w:val="{wpct}"/>' if wpct != 100 else ''
    p = f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{run("あ", fs)}{run(tcy, fs, wx + lay)}{run("いう", fs)}</w:p>'
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" '
            f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(X0 * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(Y0 * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(W * EMU)}" cy="{int(H * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="2" name="B"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
            f'<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(W * EMU)}" cy="{int(H * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
            f'<wps:txbx><w:txbxContent>{p}</w:txbxContent></wps:txbx>'
            f'<wps:bodyPr vert="eaVert" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def document(*a):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body><w:p>{box(*a)}</w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


ARMS = [(t, fs, c, w) for t in ('２４', '12', '１２３') for fs in (20, 36) for c in (0, 1) for w in (100, 88)]

if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for t, fs, c, w in ARMS:
            name = f'{ {"２４": "fw2", "12": "hw2", "１２３": "fw3"}[t]}_{fs}_c{c}_w{w}'
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/styles.xml', STYLES)
                z.writestr('word/document.xml', document(t, fs, c, w))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = fitz.open(pdf)
            ch = [(c2['c'], [round(v, 2) for v in c2['bbox']]) for b in P[0].get_text('rawdict')['blocks']
                  for l in b.get('lines', []) for s in l['spans'] for c2 in s['chars'] if c2['c'].strip()]
            print(f'{name:18} {ch}', flush=True)
    finally:
        app.Quit()
