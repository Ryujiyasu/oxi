# -*- coding: utf-8 -*-
"""Where does the body start when w:pgMar/@top is NEGATIVE?

correspondence__04a3e3e17960b59a (the JA blind floor doc, oxi 0.459) declares
  <w:pgMar w:top="-284" ... w:header="851" .../>
Oxi puts its first paragraph at y=0.  Word's first-page WordArt box, anchored
to that paragraph with posV offset 2.815pt, renders at a box top near 17pt --
i.e. Word's body appears to start near |-284|/20 = 14.2pt, not 0.

Arms sweep the top margin through negative, zero and positive values with the
header held fixed, and read Word's first body baseline out of the exported PDF.
The DELTA between arms is the answer; it cancels any font-resolution
degradation of the minimal package.

Readback: _pb_negtop_word.py (Word) / _pb_negtop_oxi.py (Oxi --dump-layout).
"""
import os
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_negtop"
os.makedirs(OUT, exist_ok=True)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
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


def header_xml(lines):
    body = "".join(
        '<w:p><w:r><w:t>H' + ("%02d" % i) + "</w:t></w:r></w:p>" for i in range(1, lines + 1)
    ) if lines else "<w:p/>"
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + body + "</w:hdr>")


def doc_xml(top_tw, header_tw, grid, bottom_tw=-284, nlines=6):
    body = "".join(
        '<w:p><w:r><w:t>L' + ("%02d" % i) + "</w:t></w:r></w:p>" for i in range(1, nlines + 1)
    )
    grid_xml = '<w:docGrid w:type="lines" w:linePitch="344"/>' if grid else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<w:body>" + body +
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>'
        '<w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
        '<w:pgMar w:top="' + str(top_tw) + '" w:right="340" w:bottom="' + str(bottom_tw) + '" w:left="340"'
        ' w:header="' + str(header_tw) + '" w:footer="510" w:gutter="0"/>'
        '<w:cols w:space="425"/>' + grid_xml +
        "</w:sectPr></w:body></w:document>"
    )


# tag -> (top_twips, header_twips, header_lines, docGrid[, bottom_twips, nlines])
ARMS = {}
for tw in (-568, -284, -142, 0, 142, 284, 851, 1134):
    ARMS["t%s" % str(tw).replace("-", "m")] = (tw, 851, 1, True)
# does the HEADER still push the body down when top is negative?
for hl in (1, 3, 6):
    ARMS["hdr%d_tm284" % hl] = (-284, 851, hl, True)
    ARMS["hdr%d_tp284" % hl] = (284, 851, hl, True)
# header distance sweep at a fixed negative top
for htw in (284, 851, 1701):
    ARMS["hd%d_tm284" % htw] = (-284, htw, 1, True)
# no docGrid control
ARMS["nogrid_tm284"] = (-284, 851, 1, False)
ARMS["nogrid_tp284"] = (284, 851, 1, False)
# bottom margin sign: 60 body lines, count how many land on page 1.
for btw in (-568, -284, 0, 284, 1134):
    ARMS["b%s" % str(btw).replace("-", "m")] = (284, 851, 1, True, btw, 60)


def build(tag):
    a = ARMS[tag]
    top_tw, header_tw, header_lines, grid = a[:4]
    bottom_tw = a[4] if len(a) > 4 else -284
    nlines = a[5] if len(a) > 5 else 6
    p = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/header1.xml", header_xml(header_lines))
        z.writestr("word/document.xml",
                   doc_xml(top_tw, header_tw, grid, bottom_tw, nlines))
    return p


if __name__ == "__main__":
    for tag in ARMS:
        print("built", build(tag), ARMS[tag])
