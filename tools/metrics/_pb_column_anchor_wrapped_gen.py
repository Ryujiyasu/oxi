# -*- coding: utf-8 -*-
"""Is relativeFrom="column" measured from a wrapped anchor paragraph's left edge?

en/policies__0026b7f7 p1: the WhatsApp text box (relativeFrom column, posOffset
-187pt, lIns 0) is drawn at x 13.9 in Word's PDF; from the margin column (54)
it would sit at -133. Its anchor paragraph wraps right of a wrapTight rectangle
spanning 1.2-192.4 (distR 9pt), i.e. its lines start at 201.4 = 13.9 + 187.

Arms: a 150 x 300pt wrapSquare rectangle (distR 9pt) anchored at the margin
column in paragraph 1, filler text wrapping beside it; a small marker
rectangle anchored with relativeFrom column / margin / page, posOffset 20pt,
in a paragraph BESIDE the big rectangle or in one BELOW it. Readout from
Word's PDF: the marker's left edge.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/colanchor'; os.makedirs(OUT, exist_ok=True)
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


def shape(sid, rel_h, x_pt, y_rel, y_pt, w_pt, h_pt, wrap, color):
    wrap_xml = '<wp:wrapSquare wrapText="bothSides"/>' if wrap else '<wp:wrapNone/>'
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{sid}" '
            f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="{rel_h}"><wp:posOffset>{int(x_pt * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="{y_rel}"><wp:posOffset>{int(y_pt * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(w_pt * EMU)}" cy="{int(h_pt * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap_xml}'
            f'<wp:docPr id="{sid}" name="S{sid}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
            f'<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(w_pt * EMU)}" cy="{int(h_pt * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>'
            f'<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


FILL = 'Filler text that wraps beside the big rectangle so the line starts to its right. '


def document(rel_h, where):
    big = shape(1, 'column', 0, 'paragraph', 0, 150, 300, True, '3366CC')
    marker = shape(2, rel_h, 20, 'paragraph', 0, 30, 15, False, 'CC3333')
    paras = [f'<w:p>{big}<w:r><w:t xml:space="preserve">{FILL}</w:t></w:r></w:p>']
    for k in range(14):
        m = marker if (where == 'beside' and k == 3) or (where == 'below' and k == 13) else ''
        paras.append(f'<w:p>{m}<w:r><w:t xml:space="preserve">Line {k}. {FILL}</w:t></w:r></w:p>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(paras)}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


import win32com.client, fitz
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
try:
    for rel_h in ('column', 'margin', 'page'):
        for where in ('beside', 'below'):
            at = os.path.abspath(os.path.join(OUT, f'{rel_h}_{where}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/styles.xml', STYLES)
                z.writestr('word/document.xml', document(rel_h, where))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = fitz.open(pdf)
            rects = []
            for dr in P[0].get_drawings():
                f = dr.get('fill')
                if f and f[0] > 0.7 and f[2] < 0.3:
                    rects.append([round(v, 2) for v in dr['rect']])
            big = [round(v, 2) for dr in P[0].get_drawings() if dr.get('fill') and dr['fill'][2] > 0.7 for v in [dr['rect'].x0, dr['rect'].x1]][:2]
            print(f'{rel_h:7} {where:7} marker {rects[:1]}  big x {big}', flush=True)
finally:
    app.Quit()
