# -*- coding: utf-8 -*-
"""How does Word paint a shape outline's compound type (a:ln@cmpd) and an
a:outerShdw effect?

ja/reports__5373dd8a: the page-frame text boxes carry a:ln w=66675
cmpd="thickThin" (Word's PDF: a 0.6w band outside, 0.2w gap, 0.2w band inside
the centre line) and the masthead roundRect an outerShdw dist=114300
dir=18900000 alpha 40% (an offset silhouette, outline included). Oxi drew a
single 5.25pt line and no shadow.

Arms: a 200 x 80pt rect text box (and a roundRect) with a 6pt outline of each
cmpd (sng, dbl, thickThin, thinThick, tri); shadows at dir 0/45/135/315 deg,
dist 9pt, blurRad 0 and 4pt, alpha 40 / 100%. Readout from Word's PDF: the
outline's filled bands (x of every edge crossing the box's mid height) and
the shadow image / fill bbox and colour.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/shape_cmpd_shadow'; os.makedirs(OUT, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
EMU = 12700
X0, Y0, W, H = 100.0, 100.0, 200.0, 80.0


def box(geom, cmpd, lw, shadow):
    ln = (f'<a:ln w="{int(lw * EMU)}" cmpd="{cmpd}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="800000"/></a:ln>'
          if lw else '<a:ln><a:noFill/></a:ln>')
    eff = ''
    if shadow:
        dirdeg, blur, alpha = shadow
        eff = (f'<a:effectLst><a:outerShdw blurRad="{int(blur * EMU)}" dist="{9 * EMU}" dir="{dirdeg * 60000}" algn="bl" rotWithShape="0">'
               f'<a:srgbClr val="000000"><a:alpha val="{alpha * 1000}"/></a:srgbClr></a:outerShdw></a:effectLst>')
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" '
            f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(X0 * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(Y0 * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(W * EMU)}" cy="{int(H * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="2" name="B"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
            f'<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(W * EMU)}" cy="{int(H * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFCC"/></a:solidFill>{ln}{eff}</wps:spPr>'
            f'<wps:txbx><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
            f'<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def document(arm):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body><w:p>{box(*arm)}</w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


ARMS = [('rect', c, 6.0, None) for c in ('sng', 'dbl', 'thickThin', 'thinThick', 'tri')]
ARMS += [('rect', 'thickThin', 3.0, None), ('roundRect', 'thickThin', 6.0, None), ('roundRect', 'dbl', 6.0, None)]
ARMS += [('rect', 'sng', 2.0, (d, b, a)) for d, b, a in ((315, 0, 40), (45, 0, 40), (135, 0, 40), (0, 0, 100), (315, 4, 40))]
ARMS += [('roundRect', 'sng', 2.0, (315, 0, 40)), ('rect', 'sng', 0, (315, 0, 40))]

if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for arm in ARMS:
            geom, cmpd, lw, sh = arm
            name = f'{geom}_{cmpd}_{lw:g}' + (f'_sh{sh[0]}_{sh[1]}_{sh[2]}' if sh else '')
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
            out = []
            for dr in P[0].get_drawings():
                r = dr['rect']
                if r.width < 50:
                    continue
                its = [(it[0], [round(v, 2) for v in it[1]] if it[0] == 're' else None) for it in dr['items']]
                out.append((dr['type'], [round(v, 2) for v in r], dr.get('width'), dr.get('fill'), dr.get('fill_opacity'),
                            [x for x in its if x[0] == 're'] or len(its)))
            ims = [([round(v, 2) for v in im['bbox']]) for im in P[0].get_image_info()]
            print(name, flush=True)
            for o in out:
                print('   ', o, flush=True)
            if ims:
                print('    images', ims, flush=True)
            if sh:
                pix = P[0].get_pixmap(dpi=144, clip=fitz.Rect(X0 - 20, Y0 - 20, X0 + W + 20, Y0 + H + 20))
                pix.save(at[:-5] + '.png')
    finally:
        app.Quit()
