# -*- coding: utf-8 -*-
"""What does relativeFrom="paragraph" reference for a float in the HEADER?

Oxi's header paint loop called resolve_floating_image_position with an EMPTY
block_y_positions, so a paragraph-anchored header float fell through to
page.margin.top -- 29.45pt below the header on an 851tw header.  The page-edge
clamp (S1268) hid it for the corpus's full-page background images by sliding
them back to y=0.

Three corpus docs say the reference is the header's own first paragraph, i.e.
`w:header`:  20f1ad3e 42.55-42.70 = -0.15, 05a78ecd 42.55-43.60 = -1.05,
1076f12a 42.55-47.20 = -4.65, all matching Word's PDF to 0.00pt.  These arms
are the self-authored repro for the same law, sweeping the one variable those
three docs hold fixed: the header distance itself.

  hd_NNN   header distance NNN twips, anchor offset -20pt
  off_NN   header 851tw, anchor offset NN pt (positive and negative)
  txt      a header that has a TEXT paragraph before the drawing's paragraph
           (does the reference follow to the second paragraph?)

Readback: _pb_hdrfloat_word.py.
"""
import os
import struct
import sys
import zlib
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_hdrfloat"
os.makedirs(OUT, exist_ok=True)
EMU = 12700.0


def png_1x1():
    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(b"\x00\x00\x00\x00")) + chunk(b"IEND", b""))


CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>"""

HRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdIMG" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/dot.png"/>
</Relationships>"""

NS = (
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'mc:Ignorable="w14"'
)

IMG_W, IMG_H = 120.0, 90.0


def anchor(voff_pt):
    cx, cy = int(round(IMG_W * EMU)), int(round(IMG_H * EMU))
    return (
        "<w:r><w:drawing>"
        '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"'
        ' relativeHeight="251658240" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>'
        + str(int(round(voff_pt * EMU)))
        + "</wp:posOffset></wp:positionV>"
        '<wp:extent cx="' + str(cx) + '" cy="' + str(cy) + '"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        '<wp:docPr id="1" name="IMG1"/><wp:cNvGraphicFramePr/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="1" name="dot.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdIMG"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + str(cx) + '" cy="' + str(cy) + '"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic>"
        "</wp:anchor></w:drawing></w:r>"
    )


def header_xml(voff, lead_text):
    ps = ""
    if lead_text:
        ps += '<w:p><w:r><w:t>HDRLEAD</w:t></w:r></w:p>'
    ps += "<w:p>" + anchor(voff) + "</w:p>"
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            "<w:hdr " + NS + ">" + ps + "</w:hdr>")


def doc_xml(header_tw):
    body = "".join('<w:p><w:r><w:t>L%02d</w:t></w:r></w:p>' % i for i in range(1, 9))
    sect = ("<w:sectPr>"
            '<w:headerReference w:type="default" r:id="rIdH"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
            ' w:header="%d" w:footer="992" w:gutter="0"/>' % header_tw
            + '<w:cols w:space="425"/>'
            "</w:sectPr>")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            "<w:document " + NS + "><w:body>" + body + sect + "</w:body></w:document>")


# tag -> (header_twips, v_posOffset_pt, lead_text_paragraph)
ARMS = {}
for _hd in (400, 851, 1200, 1600):
    ARMS["hd_%d" % _hd] = (_hd, -20.0, False)
for _off in (-40.0, -10.0, 0.0, 25.0):
    ARMS["off_%s" % str(_off).replace(".0", "").replace("-", "m")] = (851, _off, False)
ARMS["txt"] = (851, -20.0, True)


def main():
    png = png_1x1()
    for tag, (hd, voff, lead) in ARMS.items():
        path = os.path.join(OUT, tag + ".docx")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", DRELS)
            z.writestr("word/_rels/header1.xml.rels", HRELS)
            z.writestr("word/media/dot.png", png)
            z.writestr("word/header1.xml", header_xml(voff, lead))
            z.writestr("word/document.xml", doc_xml(hd))
    print("wrote %d arms to %s" % (len(ARMS), OUT))


if __name__ == "__main__":
    main()
