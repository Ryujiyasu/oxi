# -*- coding: utf-8 -*-
"""Does Word pull a floating text box back onto the page when it overflows?

layout/mod.rs `resolve_textbox_position` ends with

    if abs_x + w > page_w { abs_x = (page_w - w).max(0.0) }
    if abs_y + h > page_h { abs_y = (page_h - h).max(0.0) }

educational__0c1ced17ffb4ab6f p2 has an 855.3pt box at column+544.613 on a
1190.55pt A3 page: 598.613 + 855.3 = 1453.9 > 1190.55, so Oxi slides it to
335.25 while Word's PDF puts the box's first glyph at ~606 (= 598.613 + lIns).
The same document's `page.shapes` path (mod.rs:11909) does NOT clamp, so Oxi
draws the same rectangle twice, 263.4pt apart.

Arms sweep the amount of overflow.  Every arm is chosen so BOTH hypotheses put
the marker glyph inside the page -- absence of a glyph is not a measurement.

  hN_*   H overflow, box narrower than the page
  hP_*   same, relativeFrom="page"
  hW_*   box WIDER than the page (clamp saturates at max(0.0))
  hL_*   box starting LEFT of the page (Oxi has no clamp on that side today)
  vN_*   V overflow, relativeFrom="page"
  vW_*   box TALLER than the page
  r_*    faithful slice of educational__0c1ced17ffb4ab6f: A3 landscape, 2 cols

Readback: _pb_floatclamp_word.py (Word SaveAs2 PDF span origins).
"""
import os
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_floatclamp"
os.makedirs(OUT, exist_ok=True)
EMU = 12700.0

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:eastAsia="MS Gothic" w:cs="MS Gothic"/>
<w:sz w:val="24"/><w:szCs w:val="24"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr>
<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

NS = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'mc:Ignorable="w14 wp14"'
)

# One marker glyph at the very start of the box's only line.  The PDF origin of
# this span is  box_left + lIns  (lIns fixed at 91440 EMU = 7.2pt below), so it
# reads the resolved box origin directly.
MARK = "\u58f1"  # 壱


def anchor(w_pt, h_pt, hrel, hoff_pt, vrel, voff_pt):
    cx = int(round(w_pt * EMU))
    cy = int(round(h_pt * EMU))
    return (
        '<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps">'
        "<w:drawing>"
        '<wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0"'
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
        '<wp:docPr id="1" name="TB1"/>'
        "<wp:cNvGraphicFramePr>"
        '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
        "</wp:cNvGraphicFramePr>"
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>'
        '<wps:spPr bwMode="auto">'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="' + str(cx) + '" cy="' + str(cy) + '"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:ln w="9525"><a:solidFill><a:sysClr val="windowText" lastClr="000000"/></a:solidFill>'
        '<a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln></wps:spPr>'
        "<wps:txbx><w:txbxContent>"
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"'
        ' w:hint="eastAsia"/><w:sz w:val="24"/></w:rPr><w:t>' + MARK + "</w:t></w:r></w:p>"
        "</w:txbxContent></wps:txbx>"
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440"'
        ' bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp></a:graphicData></a:graphic>"
        "</wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>"
    )


A4 = dict(pw=11906, ph=16838, orient=None, mar=1440, cols=1, grid=False)
A3L = dict(pw=23811, ph=16838, orient="landscape", mar=1080, cols=2, grid=True)


def doc_xml(page, w_pt, h_pt, hrel, hoff, vrel, voff, lead=0):
    body = "".join('<w:p><w:r><w:t>P%02d</w:t></w:r></w:p>' % i for i in range(1, lead + 1))
    body += "<w:p>" + anchor(w_pt, h_pt, hrel, hoff, vrel, voff) + "</w:p>"
    body += "".join('<w:p><w:r><w:t>L%02d</w:t></w:r></w:p>' % i for i in range(1, 9))
    sect = (
        "<w:sectPr>"
        + '<w:pgSz w:w="%d" w:h="%d"%s/>'
        % (page["pw"], page["ph"], ' w:orient="landscape"' if page["orient"] else "")
        + '<w:pgMar w:top="1440" w:right="%d" w:bottom="1440" w:left="%d"'
        ' w:header="851" w:footer="992" w:gutter="0"/>' % (page["mar"], page["mar"])
        + ('<w:cols w:num="2" w:space="425"/>' if page["cols"] == 2
           else '<w:cols w:space="425"/>')
        + ('<w:docGrid w:type="lines" w:linePitch="360"/>' if page["grid"] else "")
        + "</w:sectPr>"
    )
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            "<w:document " + NS + "><w:body>" + body + sect + "</w:body></w:document>")


# tag -> (page, w_pt, h_pt, hrel, hoff, vrel, voff, lead)
ARMS = {}

# --- H: box 300pt wide on a 595.3pt page, column left = 72pt --------------
# left = 72 + off.  Overflow starts at off > 223.3.  Clamp predicts 295.3.
for _off in (100, 200, 224, 240, 300, 380, 450, 500):
    ARMS["hN_%d" % _off] = (A4, 300.0, 40.0, "column", float(_off), "paragraph", 0.0, 0)
# relativeFrom="page" -- same overflow, different reference
for _off in (100, 380, 500):
    ARMS["hP_%d" % _off] = (A4, 300.0, 40.0, "page", float(_off), "paragraph", 0.0, 0)
# --- box WIDER than the page: clamp saturates at 0 ------------------------
for _off in (0, 60, 200):
    ARMS["hW_%d" % _off] = (A4, 700.0, 40.0, "column", float(_off), "paragraph", 0.0, 0)
# --- box hanging off the LEFT edge (Oxi never clamps this side) -----------
for _off in (-40, -70, -120):
    ARMS["hL_m%d" % abs(_off)] = (A4, 300.0, 40.0, "column", float(_off), "paragraph", 0.0, 0)

# --- V: box 100pt tall, page 841.9pt, relativeFrom="page" ----------------
# Overflow starts at off > 741.9.  Clamp predicts 741.9.
for _off in (600, 742, 760, 800):
    ARMS["vN_%d" % _off] = (A4, 200.0, 100.0, "column", 100.0, "page", float(_off), 0)
# --- box TALLER than the page: clamp saturates at 0 ----------------------
for _off in (100, 300):
    ARMS["vW_%d" % _off] = (A4, 200.0, 900.0, "column", 100.0, "page", float(_off), 0)

# --- faithful slice of the corpus doc ------------------------------------
# A3 landscape, 2 cols, 855.3x33.5 box at column + 544.613.
ARMS["r_edu"] = (A3L, 855.3, 33.5, "column", 544.613, "paragraph", 19.75, 0)
# same box anchored from a paragraph that flows in COLUMN 2 (lead fills col 1)
ARMS["r_edu_c2"] = (A3L, 855.3, 33.5, "column", 544.613, "paragraph", 19.75, 40)


def main():
    for tag, (page, w, h, hrel, hoff, vrel, voff, lead) in ARMS.items():
        path = os.path.join(OUT, tag + ".docx")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", DRELS)
            z.writestr("word/styles.xml", STYLES)
            z.writestr("word/document.xml", doc_xml(page, w, h, hrel, hoff, vrel, voff, lead))
    print("wrote %d arms to %s" % (len(ARMS), OUT))


if __name__ == "__main__":
    main()
