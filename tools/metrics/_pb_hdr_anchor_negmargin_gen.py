# -*- coding: utf-8 -*-
"""Where does Word put a HEADER picture anchored relativeFrom="margin" with a
NEGATIVE offset, and does the body flow below it?

en/administrative__00777f39 (compat 15, top margin 72, header 36): a 95.25pt
logo in the header, positionV margin -685800 EMU (-54pt), wrapTopAndBottom,
allowOverlap 0. Word draws it at y 73.8 (not 18) and the body starts under it.

Arms: offset -54 / 0 / +18pt x wrap topAndBottom / none / square x relativeFrom
margin / page / paragraph x allowOverlap 0 / 1.
Readout: Word PDF image bbox and the first body baseline.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/hdr_anchor_negmargin'; os.makedirs(OUT, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
HRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def png():
    import struct, zlib
    w = h = 8
    raw = b''.join(b'\x00' + b'\x40\x40\x40' * w for _ in range(h))
    def chunk(t, d):
        c = struct.pack('>I', len(d)) + t + d
        return c + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))


def header(off_pt, wrap, rel, overlap):
    wraps = {'tb': '<wp:wrapTopAndBottom/>', 'none': '<wp:wrapNone/>', 'sq': '<wp:wrapSquare wrapText="bothSides"/>'}
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}><w:p><w:r><w:drawing>'
            f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="{overlap}">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>2400300</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="{rel}"><wp:posOffset>{int(off_pt * 12700)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="1247775" cy="1209675"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wraps[wrap]}<wp:docPr id="1" name="P"/>'
            '<wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            '<pic:nvPicPr><pic:cNvPr id="0" name="P"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1247775" cy="1209675"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:hdr>')


def document():
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
            '<w:p><w:r><w:t>BODYLINE one</w:t></w:r></w:p><w:p><w:r><w:t>BODYLINE two</w:t></w:r></w:p>'
            '<w:sectPr><w:headerReference w:type="default" r:id="rId8"/><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for off, wrap, rel, ov in itertools.product((-54, 0, 18), ('tb', 'none', 'sq'), ('margin', 'page', 'paragraph'), (0, 1)):
            name = f'o{off}_{wrap}_{rel}_ov{ov}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', SETTINGS)
                z.writestr('word/_rels/header1.xml.rels', HRELS); z.writestr('word/media/image1.png', png())
                z.writestr('word/header1.xml', header(off, wrap, rel, ov)); z.writestr('word/document.xml', document())
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            img = [round(v, 1) for i in P.get_image_info() for v in i['bbox']][:4]
            body = [round(l['spans'][0]['origin'][1], 1) for b in P.get_text('dict')['blocks'] for l in b.get('lines', [])
                    if l['spans'][0]['text'].startswith('BODYLINE')]
            print(f'{name:28} img {img}  body {body}', flush=True)
    finally:
        app.Quit()
