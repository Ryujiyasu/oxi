# -*- coding: utf-8 -*-
"""Where does Word stop drawing a text box's overflowing lines?

ja/reference__22761439: two 250 x 660pt column text boxes hold ~53 lines of 18pt
exact «○○○» each; Word draws them past the box bottom but its last drawn
baselines are 804 / 811 (bottom margin at 813.5, footer at 28.35); Oxi draws
every line down to 1100 (S1644 lets text run past the box bottom).

Arms: a 200 x 200pt box at page y 500 holding 40 lines of 18pt exact text,
page bottom margin 72 / 144pt, footer distance 36pt, with and without a footer
line; plus a box at y 100 (overflow ends mid-page). Readout: the last drawn
line's baseline in Word's PDF.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from _pb_shape_cmpd_shadow_gen import NS, CT, RR, STYLES, EMU
OUT = 'tests/fixtures/txbx_overflow_clip'; os.makedirs(OUT, exist_ok=True)
CT2 = CT.replace('</Types>', '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
FTR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
       '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>{t}</w:t></w:r></w:p></w:ftr>')


def box(y, h, n, vo='', bins=3.6):
    paras = ''.join(f'<w:p><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="exact"/></w:pPr><w:r><w:t>Line{k:02d}</w:t></w:r></w:p>' for k in range(n))
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" '
            f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{100 * EMU}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(y * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{200 * EMU}" cy="{int(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="2" name="B"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
            f'<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{200 * EMU}" cy="{int(h * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
            f'<wps:txbx><w:txbxContent>{paras}</w:txbxContent></wps:txbx>'
            f'<wps:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="{int(bins * EMU)}" anchor="t"{vo}/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def document(y, h, n, bottom, footer, vo='', bins=3.6):
    ref = '<w:footerReference w:type="default" r:id="rId7"/>' if footer else ''
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body><w:p>{box(y, h, n, vo, bins)}</w:p>'
            f'<w:sectPr>{ref}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="{int(bottom * 20)}" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


ARMS = [(500, 200, 40, 72, False), (500, 200, 40, 144, False), (500, 200, 40, 72, True), (500, 200, 40, 144, True),
        (100, 200, 20, 72, False), (600, 100, 30, 36, False)]
if os.environ.get('PB_SET') == '2':
    ARMS = [(600, 100, 30, 36, False, v) for v in (' vertOverflow="overflow"', ' vertOverflow="clip"', ' vertOverflow="ellipsis"')]
    ARMS += [(600, h, 30, 36, False, '') for h in (95, 98, 101, 104, 107)]
if os.environ.get('PB_SET') == '3':
    ARMS = [(600, h, 30, 36, False, '', b) for b in (0.0, 10.0) for h in range(92, 122, 2)]

if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for arm in ARMS:
            y, h, n, bottom, footer = arm[:5]
            vo = arm[5] if len(arm) > 5 else ''
            bins = arm[6] if len(arm) > 6 else 3.6
            name = f'y{y}_h{h}_n{n}_bm{bottom}_{"ftr" if footer else "noftr"}' + (vo.split('"')[1] if vo else '') + (f'_b{bins:g}' if len(arm) > 6 else '')
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT2); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/styles.xml', STYLES)
                z.writestr('word/footer1.xml', FTR.format(t='Footer line')); z.writestr('word/document.xml', document(y, h, n, bottom, footer, vo, bins))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = fitz.open(pdf)
            ys = [(s['text'], round(s['origin'][1], 2)) for b in P[0].get_text('dict')['blocks'] for l in b.get('lines', [])
                  for s in l['spans'] if s['text'].startswith('Line')]
            print(f'{name:28} pages {len(P)} drawn {len(ys)} last {ys[-1] if ys else None}', flush=True)
    finally:
        app.Quit()
