# -*- coding: utf-8 -*-
"""Same question as _pb_floatclamp_gen, for a floating IMAGE.

`resolve_floating_image_position` clamps only the Y axis:

    if abs_y + img.height > page_h { abs_y = (page_h - img.height).max(0.0) }

The text-box arms showed Word clamps neither axis.  An image is a different
object, so it gets its own measurement rather than an inherited rule.

The image is a 2x2 PNG scaled to the arm's box, with a black pixel in its
TOP-LEFT quadrant only, so the drawn rect can be read off the rendered page
by ink instead of by a glyph origin.
"""
import os
import struct
import sys
import zlib
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_floatclamp_img"
os.makedirs(OUT, exist_ok=True)
EMU = 12700.0


def png_2x2():
    """2x2 RGB PNG: black top-left, white elsewhere."""
    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", 2, 2, 8, 2, 0, 0, 0)
    W, B = b"\xff\xff\xff", b"\x00\x00\x00"
    raw = b"\x00" + B + W + b"\x00" + W + W
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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


def anchor(w_pt, h_pt, hrel, hoff_pt, vrel, voff_pt):
    cx = int(round(w_pt * EMU))
    cy = int(round(h_pt * EMU))
    return (
        "<w:r><w:drawing>"
        '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"'
        ' relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="' + hrel + '"><wp:posOffset>'
        + str(int(round(hoff_pt * EMU)))
        + "</wp:posOffset></wp:positionH>"
        '<wp:positionV relativeFrom="' + vrel + '"><wp:posOffset>'
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


def doc_xml(w_pt, h_pt, hrel, hoff, vrel, voff):
    body = "<w:p>" + anchor(w_pt, h_pt, hrel, hoff, vrel, voff) + "</w:p>"
    body += "".join('<w:p><w:r><w:t>L%02d</w:t></w:r></w:p>' % i for i in range(1, 9))
    sect = ("<w:sectPr>"
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
            ' w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="425"/>'
            "</w:sectPr>")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            "<w:document " + NS + "><w:body>" + body + sect + "</w:body></w:document>")


# tag -> (w_pt, h_pt, hrel, hoff, vrel, voff)
ARMS = {}
# V overflow: image 100pt tall, page 841.9.  Clamp predicts top 741.9.
for _off in (600, 760, 800):
    ARMS["ivN_%d" % _off] = (200.0, 100.0, "column", 100.0, "page", float(_off))
# image TALLER than the page: clamp saturates at 0.
ARMS["ivW_200"] = (200.0, 900.0, "column", 100.0, "page", 200.0)
# H overflow (Oxi does not clamp this axis today -- confirm Word agrees).
for _off in (380, 500):
    ARMS["ihN_%d" % _off] = (300.0, 60.0, "column", float(_off), "page", 200.0)


def main():
    png = png_2x2()
    for tag, (w, h, hrel, hoff, vrel, voff) in ARMS.items():
        path = os.path.join(OUT, tag + ".docx")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", DRELS)
            z.writestr("word/media/dot.png", png)
            z.writestr("word/document.xml", doc_xml(w, h, hrel, hoff, vrel, voff))
    print("wrote %d arms to %s" % (len(ARMS), OUT))


if __name__ == "__main__":
    main()
