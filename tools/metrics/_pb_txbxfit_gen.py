# -*- coding: utf-8 -*-
"""How many lines does Word draw in a FIXED-HEIGHT floating text box?

Oxi drops every line whose index >= floor(inner_h / line_h) (layout/mod.rs
`line_cutoff_y`).  When the box is shorter than ONE line that floor is 0 and
the box's ENTIRE text vanishes -- 61 boxes across 17 of the 200 JA corpus
docs, including the blind-set floor doc correspondence__04a3e3e17960b59a
(12 boxes) and policies__04a2204e7119e51d (the LO-win #1).

This probe pins Word's actual rule.  One floating wps text box (wrapNone,
AlternateContent inside a run -- the shape 51/61 of the corpus boxes have),
MS Gothic 12pt, docGrid lines/360 like the real docs.  Only the box HEIGHT
and the LINE COUNT vary.

Arms:
  a1_hNN    1 paragraph, box height NN pt  (does the single line survive?)
  b3_hNN    3 paragraphs, box height NN pt (how many of the 3 survive?)
  c3_hNN    same as b3 but bodyPr vertOverflow="clip"
  n3_hNN    same as b3 but NO docGrid (is the grid snap the whole story?)

Readback: _pb_txbxfit_word.py (Word SaveAs2 PDF) / _pb_txbxfit_oxi.py.
"""
import os
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_txbxfit"
os.makedirs(OUT, exist_ok=True)

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

EMU = 12700.0  # per point
# The marker glyphs are unique per line so the PDF readback can tell which
# lines Word kept.  Kanji numerals: they exist in MS Gothic and nowhere else
# in the probe.
MARKS = ["\u58f1", "\u5f10", "\u53c2"]  # 壱 弐 参
# Latin markers for arms that must NOT trigger Word CJK font fallback
# (a kanji in an "Arial" run silently resolves to a CJK face).
LMARKS = ["Q", "W", "E"]


def txbx_paragraphs(n, sz=24, face="MS Gothic", line=None, latin=False):
    out = []
    for i in range(n):
        out.append(
            '<w:p><w:pPr><w:spacing w:after="0" ' +
            ('w:line="240" w:lineRule="auto"' if line is None
             else 'w:line="' + str(line) + '" w:lineRule="exact"') + '/>'
            '<w:rPr><w:rFonts w:ascii="' + face + '" w:eastAsia="' + face + '" w:hAnsi="' + face + '"/>'
            '<w:sz w:val="' + str(sz) + '"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="' + face + '" w:eastAsia="' + face + '" w:hAnsi="' + face + '"' +
            ('' if latin else ' w:hint="eastAsia"') + '/>'
            '<w:sz w:val="' + str(sz) + '"/></w:rPr><w:t>' +
            (LMARKS if latin else MARKS)[i] + "</w:t></w:r></w:p>"
        )
    return "".join(out)


def anchor(height_pt, nlines, overflow, sz=24, face="MS Gothic", tins=45720, bins=45720, line=None, latin=False):
    cx = int(round(140.0 * EMU))
    cy = int(round(height_pt * EMU))
    ovf = ' vertOverflow="clip"' if overflow else ""
    return (
        '<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps">'
        "<w:drawing>"
        '<wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0"'
        ' relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>2540000</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
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
        "<a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a:ln></wps:spPr>"
        "<wps:txbx><w:txbxContent>" + txbx_paragraphs(nlines, sz, face, line, latin) + "</w:txbxContent></wps:txbx>"
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="' + str(tins) + '" rIns="91440"'
        ' bIns="' + str(bins) + '" anchor="t" anchorCtr="0"' + ovf + "><a:noAutofit/></wps:bodyPr>"
        "</wps:wsp></a:graphicData></a:graphic>"
        "</wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>"
    )


def doc_xml(height_pt, nlines, overflow, grid, sz=24, face="MS Gothic", tins=45720, bins=45720, line=None, latin=False):
    # Body: the anchor paragraph, then numbered body lines so the readback can
    # locate the box relative to the flow.
    body = "<w:p>" + anchor(height_pt, nlines, overflow, sz, face, tins, bins, line, latin) + "</w:p>"
    body += "".join(
        '<w:p><w:r><w:t>L' + ("%02d" % i) + "</w:t></w:r></w:p>" for i in range(1, 9)
    )
    grid_xml = '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        "<w:document " + NS + "><w:body>" + body +
        "<w:sectPr>"
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        "<w:cols w:space=\"425\"/>" + grid_xml +
        "</w:sectPr></w:body></w:document>"
    )


# (tag) -> (height_pt, nlines, vertOverflow_clip, docGrid)
# EXTRA[tag] -> (sz_halfpoints, face, tIns_emu, bIns_emu); default (24, MS Gothic, 45720, 45720)
ARMS = {}
EXTRA = {}
for h in (10, 14, 18, 21, 23.45, 26, 30):
    ARMS["a1_h%s" % str(h).replace(".", "p")] = (h, 1, False, True)
for h in (20, 30, 40, 46, 58, 70):
    ARMS["b3_h%d" % h] = (h, 3, False, True)
for h in (20, 40, 58):
    ARMS["c3_h%d" % h] = (h, 3, True, True)
for h in (10, 16, 20, 40, 58):
    ARMS["n3_h%d" % h] = (h, 3 if h >= 20 else 1, False, False)

# --- discriminators for "how much height does ONE line demand?" ---------
# Two hypotheses survive the arms above (both fit all 21):
#   A  required(n) = 2*tIns + bIns + (n-1)*L          (insets twice over)
#   B  required(n) = font_size    + (n-1)*L           (the em box, insets ignored)
# e1 pins required(1) to +-0.25pt.  e2/e3 move the insets while holding the
# font: A tracks them, B does not.  e4/e5 move the font: B tracks it, A does not.
def _t(h):
    return str(h).replace(".", "p").replace("-", "m")

for h in (10.5, 11, 11.5, 12, 12.5, 13):
    ARMS["e1_h" + _t(h)] = (h, 1, False, False)
for h in (2, 6, 11, 12, 13):
    tag = "e2_ins0_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "MS Gothic", 0, 0)
for h in (12, 13, 25, 31, 33):
    tag = "e3_ins14_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "MS Gothic", 182880, 45720)
for h in (14, 17, 18, 19, 22):
    tag = "e4_fs18_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (36, "MS Gothic", 45720, 45720)
for h in (11, 12, 13, 13.5, 14):
    tag = "e5_arial_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "Arial", 45720, 45720)

# f: insets ZEROED, fine height sweep, three (face, size) pairs.  With
# tIns=bIns=0 the threshold IS the unknown constant X in
#     required(n) = tIns + bIns + (n-1)*L + X
# e1/e2/e3 bracket X to (3.3, 3.8] at MS Gothic 12pt.  Candidates that fit:
#   fixed 3.6pt (0.05in)  |  0.3 * font_size  |  a font descent term.
# The three pairs below separate them: 0.3*fs predicts 3.6 / 5.4 / 2.4,
# a fixed constant predicts 3.6 / 3.6 / 3.6.
for h in (2, 3, 3.4, 3.6, 3.8, 4.2, 5):
    tag = "f12_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "MS Gothic", 0, 0)
for h in (3, 3.6, 4.2, 5, 5.4, 6, 7):
    tag = "f18_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (36, "MS Gothic", 0, 0)
for h in (1.8, 2.2, 2.4, 2.6, 3, 3.6, 4):
    tag = "f08_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (16, "MS Gothic", 0, 0)
for h in (2, 3, 3.4, 3.6, 3.8, 4.2, 5):
    tag = "fA12_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "Arial", 0, 0)

# g: tall box, insets 0, TWO lines -> reads A (first baseline - box top) and
# L (baseline advance) for each (face, size), so X can be tested against them.
for sz, face in ((16, "MS Gothic"), (24, "MS Gothic"), (36, "MS Gothic"),
                 (24, "Arial"), (36, "Arial")):
    tag = "g_%s%d" % ("msg" if face == "MS Gothic" else "ari", sz)
    ARMS[tag] = (200, 2, False, False)
    EXTRA[tag] = (sz, face, 0, 0)

# h: does X follow the FONT SIZE or the LINE HEIGHT?  Same 12pt font, line
# spacing forced EXACT to 30pt / 8pt.  X constant -> font; X moves -> line.
for h in (3, 3.4, 3.6, 4, 6, 10):
    tag = "h30_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "MS Gothic", 0, 0, 600)
for h in (1.6, 2, 2.4, 2.8, 3.2, 3.6):
    tag = "h08_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "MS Gothic", 0, 0, 160)
# finer 18pt, plus 24pt and 10pt, all insets 0
for h in (4.2, 4.4, 4.6, 4.8, 5.0):
    tag = "i18_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (36, "MS Gothic", 0, 0)
for h in (5, 5.5, 6, 6.5, 7, 7.5):
    tag = "i24_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (48, "MS Gothic", 0, 0)
for h in (2.6, 2.8, 3.0, 3.2, 3.4):
    tag = "i10_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (20, "MS Gothic", 0, 0)
for tag, sz in (("g_msg48", 48), ("g_msg20", 20)):
    ARMS[tag] = (200, 2, False, False)
    EXTRA[tag] = (sz, "MS Gothic", 0, 0)
for tag, line in (("g_ln30", 600), ("g_ln08", 160)):
    ARMS[tag] = (200, 2, False, False)
    EXTRA[tag] = (24, "MS Gothic", 0, 0, line)

# j: a REAL Latin face (Latin markers, no eastAsia hint).  Is the allowance
# c in  "n*L <= inner_h + c"  a font metric or a size-only Word constant?
for h in (2.8, 3.0, 3.2, 3.4, 3.6, 3.8, 4.2, 5.0):
    tag = "j12_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "Arial", 0, 0, None, True)
for h in (5.2, 5.6, 6.0, 6.4, 6.8, 7.2, 8.0, 9.0):
    tag = "j24_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (48, "Arial", 0, 0, None, True)
for h in (3.2, 3.4, 3.6, 3.8, 4.0, 4.4, 5.0):
    tag = "jT12_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "Times New Roman", 0, 0, None, True)
for tag, sz, face in (("g_j12", 24, "Arial"), ("g_j24", 48, "Arial"),
                      ("g_jT12", 24, "Times New Roman")):
    ARMS[tag] = (200, 2, False, False)
    EXTRA[tag] = (sz, face, 0, 0, None, True)

# k: the exact-line-spacing outlier (h30 dropped even at H=10 while the
# fs-only reading predicts ~3.5) and a face with a very different L/fs
# (Meiryo ~1.6em) to separate "X follows the font size" from "X follows the
# line box".
for h in (12, 15, 18, 20, 22, 24, 26, 30):
    tag = "k30_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "MS Gothic", 0, 0, 600)
for h in (2.6, 3.0, 3.4, 3.8, 4.2, 4.6, 5.0):
    tag = "kMe12_h" + _t(h)
    ARMS[tag] = (h, 1, False, False)
    EXTRA[tag] = (24, "Meiryo", 0, 0)
for tag, sz, face, line in (("g_me12", 24, "Meiryo", None),):
    ARMS[tag] = (200, 2, False, False)
    EXTRA[tag] = (sz, face, 0, 0, line)

# m: does a text-box line snap to a MULTIPLE of the docGrid pitch?  Oxi does
# (ceil(natural / pitch) * pitch), which turns a 14pt CJK line (natural 18.2)
# in an 18pt grid into a 36pt line -- that alone empties every box shorter
# than 36pt.  Tall box, 2 lines, grid lines/360 (18pt), size swept across the
# pitch boundary: the baseline ADVANCE is Word's line height.
for sz in (24, 26, 28, 32, 36, 40, 48):
    tag = "m_grid18_sz%d" % sz
    ARMS[tag] = (200, 2, False, True)
    EXTRA[tag] = (sz, "MS Gothic", 45720, 45720)
# same sizes with NO grid, as the natural-line control
for sz in (24, 26, 28, 32, 36, 40, 48):
    tag = "m_nogrid_sz%d" % sz
    ARMS[tag] = (200, 2, False, False)
    EXTRA[tag] = (sz, "MS Gothic", 45720, 45720)








def build(tag):
    height_pt, nlines, overflow, grid = ARMS[tag]
    ex = EXTRA.get(tag, (24, "MS Gothic", 45720, 45720))
    sz, face, tins, bins = ex[:4]
    line = ex[4] if len(ex) > 4 else None
    latin = ex[5] if len(ex) > 5 else False
    p = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml",
                   doc_xml(height_pt, nlines, overflow, grid, sz, face, tins, bins, line, latin))
    return p


if __name__ == "__main__":
    for tag in ARMS:
        print("built", build(tag), ARMS[tag])
