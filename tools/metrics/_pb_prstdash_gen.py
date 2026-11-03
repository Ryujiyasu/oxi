# -*- coding: utf-8 -*-
"""Word's dash arrays for every a:prstDash on a text-box outline (OXI_TXBX_DASH
took them from ECMA-376 20.1.10.48 after measuring only dashDot).
golden tsuuchisho_002669443 (lgDashDotDot 0.5pt) regressed -0.0016.
Arms: a 200 x 80 rect box, outline 0.5 / 1.5pt, each preset. Readout: the
PDF stroke's dash array and phase."""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from _pb_shape_cmpd_shadow_gen import NS, CT, RR, DRELS, STYLES, EMU, X0, Y0, W, H
OUT = 'tests/fixtures/prstdash'; os.makedirs(OUT, exist_ok=True)
PRESETS = ['dash', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot', 'dot']
def box(p, lw):
    ln = f'<a:ln w="{int(lw * EMU)}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="{p}"/></a:ln>'
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" '
            f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(X0 * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(Y0 * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(W * EMU)}" cy="{int(H * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="2" name="B"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
            f'<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(W * EMU)}" cy="{int(H * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>{ln}</wps:spPr>'
            f'<wps:txbx><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
            f'<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for lw in (0.5, 1.5):
            for p in PRESETS:
                at = os.path.abspath(os.path.join(OUT, f'{p}_{lw:g}.docx'))
                doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body><w:p>{box(p, lw)}</w:p>'
                       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
                       'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
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
                dr = [(x.get('dashes'), x.get('width'), x.get('lineCap'), len(x['items']), [round(v, 2) for v in x['rect']]) for x in P[0].get_drawings() if x['rect'].width > 100]
                print(f'{p:14} {lw:g} {dr}', flush=True)
    finally:
        app.Quit()
