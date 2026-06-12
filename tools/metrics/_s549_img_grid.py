# -*- coding: utf-8 -*-
"""S549 — inline image-only paragraph height under docGrid lines.
Hypothesis (3a4f live data): block = ceil(extent/pitch)*pitch (185@18 -> 198).
Matrix: extent {185, 100, 90, 36} x grid {lines360, none}.
Measures Y of the para before vs after the image para (3 paras: 国/img/国).
Needs a real image: 1x1 PNG scaled via extent. cp932-safe.
"""
import base64
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s549_imggrid')
os.makedirs(OUT, exist_ok=True)

PNG = base64.b64decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==')

CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>')
STYLES = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
          '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')

EMU = 914400  # per inch; pt -> emu = pt/72*914400 = pt*12700


def doc_xml(extent_pt, grid):
    cx = int(300 * 12700)
    cy = int(extent_pt * 12700)
    drawing = ('<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" '
               'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
               '<wp:extent cx="%d" cy="%d"/><wp:docPr id="1" name="P1"/>'
               '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
               '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
               '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
               '<pic:nvPicPr><pic:cNvPr id="1" name="P1"/><pic:cNvPicPr/></pic:nvPicPr>'
               '<pic:blipFill><a:blip r:embed="rId2" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
               '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
               '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
               '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
               '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
               % (cx, cy, cx, cy))
    body = ('<w:p><w:r><w:t>国国国</w:t></w:r></w:p>'
            '<w:p><w:r>%s</w:r></w:p>'
            '<w:p><w:r><w:t>国国国</w:t></w:r></w:p>' % drawing)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"/>%s</w:sectPr>'
            % ('<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ''))
    return ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body>%s%s</w:body></w:document>') % (body, sect)


def build(docx, extent_pt, grid):
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc_xml(extent_pt, grid))
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/media/image1.png', PNG)


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for grid in (True, False):
        for ext in (185, 100, 90, 36):
            tag = 's549_%s_e%d' % ('grid' if grid else 'none', ext)
            docx = os.path.join(OUT, tag + '.docx')
            build(docx, ext, grid)
            wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
            try:
                ys = []
                for p in list(wdoc.Paragraphs)[:3]:
                    s = p.Range.Start
                    ys.append(wdoc.Range(s, s).Information(6))
                blk = ys[2] - ys[1]
                print('%s: y1=%.2f y_img=%.2f y3=%.2f -> img block=%.2f (extent %d, /18=%.2f)'
                      % (tag, ys[0], ys[1], ys[2], blk, ext, blk / 18.0))
            finally:
                wdoc.Close(False)
finally:
    word.Quit()
