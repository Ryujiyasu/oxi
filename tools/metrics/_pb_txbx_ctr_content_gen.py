# -*- coding: utf-8 -*-
"""What height does Word centre in an anchor="ctr" text box: the ink, or the
laid-out paragraphs (empty ones and spacing included)? And an overflowing
content -- top-aligned or centred upward?

ja/reports__5373dd8a p2: the 688.5pt frame (anchor ctr) ends its text with 34
empty centred paragraphs, far past the box; Word puts the first baseline at
box top + 3.6 inset + 2.625 (half the 5.25pt outline) + the grid line's own
offset -- i.e. top-aligned. Oxi centred the ink (the empties draw nothing)
and set the text 4.9pt low.

Arms (200 x 200pt box, anchor ctr, Calibri 11 single): 3 text lines; + 2 / + 20
trailing empty paragraphs; + 2 leading empty paragraphs; last paragraph
spacing after 24pt; first paragraph spacing before 24pt; 20 text lines
(overflow). Readout: the first text line's baseline in Word's PDF.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from _pb_shape_cmpd_shadow_gen import NS, CT, RR, DRELS, STYLES, EMU
CLNSP = ' compatLnSpc="1"' if os.environ.get('PB_CLNSP') else ''
OUT = 'tests/fixtures/txbx_ctr_content' + ('_clnsp' if CLNSP else ''); os.makedirs(OUT, exist_ok=True)
X0, Y0, W, H = 100.0, 100.0, 200.0, 200.0


def para(t, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr>' + (f'<w:r><w:t>{t}</w:t></w:r>' if t else '') + '</w:p>'


def content(kind):
    text = [para(f'Line {k}') for k in range(3)]
    if kind == 'plain':
        return ''.join(text)
    if kind.startswith('trail'):
        return ''.join(text) + para('') * int(kind[5:])
    if kind == 'lead2':
        return para('') * 2 + ''.join(text)
    if kind == 'after24':
        return ''.join(text[:2]) + para('Line 2', '<w:spacing w:after="480"/>')
    if kind == 'before24':
        return para('Line 0', '<w:spacing w:before="480"/>') + ''.join(text[1:])
    if kind == 'over20':
        return ''.join(para(f'Line {k}') for k in range(20))


def box(kind):
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" '
            f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(X0 * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(Y0 * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(W * EMU)}" cy="{int(H * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="2" name="B"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
            f'<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(W * EMU)}" cy="{int(H * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
            f'<wps:txbx><w:txbxContent>{content(kind)}</w:txbxContent></wps:txbx>'
            f'<wps:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"{CLNSP}/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def document(kind):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body><w:p>{box(kind)}</w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


ARMS = ['plain', 'trail2', 'trail20', 'lead2', 'after24', 'before24', 'over20']

if __name__ == '__main__' and os.environ.get('PB_SET') != '2':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for kind in ARMS:
            at = os.path.abspath(os.path.join(OUT, f'{kind}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/styles.xml', STYLES)
                z.writestr('word/document.xml', document(kind))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = fitz.open(pdf)
            ys = [round(s['origin'][1], 2) for b in P[0].get_text('dict')['blocks'] for l in b.get('lines', [])
                  for s in l['spans'] if s['text'].startswith('Line')]
            print(f'{kind:10} first baseline {ys[:1]} last {ys[-1:]} n {len(ys)}', flush=True)
    finally:
        app.Quit()


# Second set (run with PB_SET=2): text that fits but whose last space-after
# overflows the box -- en/policies__009625de «Our Vision and Values» (Comic Sans
# MS bold 28, line 276 auto, after 200tw, 372 x 57pt box, anchor ctr).
def box2(after_tw, h_pt, line):
    p = (f'<w:p><w:pPr><w:spacing w:after="{after_tw}" w:line="{line}" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'
         '<w:r><w:rPr><w:rFonts w:ascii="Comic Sans MS" w:hAnsi="Comic Sans MS"/><w:b/><w:sz w:val="56"/></w:rPr><w:t>Our Vision and Values</w:t></w:r></w:p>')
    return box('plain').replace(content('plain'), p).replace(f'cy="{int(H * EMU)}"', f'cy="{int(h_pt * EMU)}"')


if __name__ == '__main__' and os.environ.get('PB_SET') == '2':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for after_tw in (0, 100, 200, 400):
            for h_pt in (57.0, 70.0, 100.0):
                for line in (240, 276):
                    name = f'set2_a{after_tw}_h{h_pt:g}_l{line}'
                    at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
                    doc = document('plain').replace(box('plain'), box2(after_tw, h_pt, line))
                    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                        z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/styles.xml', STYLES)
                        z.writestr('word/document.xml', doc)
                    pdf = at[:-5] + '.pdf'
                    d = app.Documents.Open(at, ReadOnly=True)
                    try:
                        d.ExportAsFixedFormat(pdf, 17)
                    finally:
                        d.Close(False)
                    P = fitz.open(pdf)
                    ys = [round(s['origin'][1], 2) for b in P[0].get_text('dict')['blocks'] for l in b.get('lines', [])
                          for s in l['spans'] if s['text'].startswith('Our')]
                    print(f'{name:24} baseline {ys}', flush=True)
    finally:
        app.Quit()
