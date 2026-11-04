# -*- coding: utf-8 -*-
"""Centred text box whose one line fits but whose space-after does not, under
compat 14 + compatLnSpc and multiple line spacing.

en/policies__009625de «Our Vision and Values»: Comic Sans MS bold 28, line 276
auto, after 200tw, a 372 x 57pt box, anchor ctr, compatLnSpc=1, compat 14.
Word's baseline is 33.92 under the line top Oxi uses when it top-aligns.

Arms: after 0 / 100 / 200 / 400tw x line 240 / 276 x box height 57 / 80 /
120pt x compat 14 / 15 (compatLnSpc on). Readout: the baseline.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from _pb_shape_cmpd_shadow_gen import NS, CT, RR, STYLES, EMU
OUT = 'tests/fixtures/txbx_ctr_clnsp'; os.makedirs(OUT, exist_ok=True)
CT2 = CT.replace('</Types>', '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
X0, Y0, W = 100.0, 100.0, 372.0


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


def document(after_tw, line, h):
    p = (f'<w:p><w:pPr><w:spacing w:after="{after_tw}" w:line="{line}" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'
         '<w:r><w:rPr><w:rFonts w:ascii="Comic Sans MS" w:hAnsi="Comic Sans MS"/><w:b/><w:sz w:val="56"/></w:rPr><w:t>Our Vision and Values</w:t></w:r></w:p>')
    bx = (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" '
          f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
          f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(X0 * EMU)}</wp:posOffset></wp:positionH>'
          f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(Y0 * EMU)}</wp:posOffset></wp:positionV>'
          f'<wp:extent cx="{int(W * EMU)}" cy="{int(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
          f'<wp:docPr id="2" name="B"/><wp:cNvGraphicFramePr/>'
          f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
          f'<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(W * EMU)}" cy="{int(h * EMU)}"/></a:xfrm>'
          f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
          f'<wps:txbx><w:txbxContent>{p}</w:txbxContent></wps:txbx>'
          f'<wps:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr" compatLnSpc="1"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body><w:p>{bx}</w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


ARMS = [(a, l, h, c) for c in (14, 15) for l in (240, 276) for h in (57.0, 80.0, 120.0) for a in (0, 100, 200, 400)]

if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for a, l, h, c in ARMS:
            name = f'c{c}_l{l}_h{h:g}_a{a}'
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT2); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/styles.xml', STYLES)
                z.writestr('word/settings.xml', settings(c)); z.writestr('word/document.xml', document(a, l, h))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = fitz.open(pdf)
            ys = [round(s['origin'][1], 2) for b in P[0].get_text('dict')['blocks'] for ln in b.get('lines', [])
                  for s in ln['spans'] if s['text'].startswith('Our')]
            print(f'{name:22} baseline {ys} (- inner top {round(ys[0] - Y0 - 3.6, 2) if ys else None})', flush=True)
    finally:
        app.Quit()
