# -*- coding: utf-8 -*-
"""Where does Word put the BASELINE inside an `exact` line box?

`probexexactclip_exactclip` is the whole golden corpus's SSIM floor (0.4638).
Its line pitch is already right in Oxi -- both sides step 8.00pt -- and the ink
sits a flat **2.40pt too low** on all 3 pages (measured by cross-correlating the
row-ink profiles of Word's own PDF against Oxi's render, so no box-vs-ink
convention is involved). So the defect is purely WHERE IN THE BOX the glyphs go
when `w:lineRule="exact"` makes the box SMALLER than the text.

One document gives one point (Word puts the baseline 6.38pt below the box top
for a 10.5pt MS Mincho line in an 8.00pt box). Sweep it: the exact line value
from far below the natural height to above it, at three sizes and two faces,
and read the baseline straight out of Word's PDF (`span["origin"][1]`, exact).

    python tools/metrics/_pb_exactbase_gen.py         # build
    python tools/metrics/_pb_exactbase_read.py word   # Word truth (COM -> PDF)
    python tools/metrics/_pb_exactbase_read.py oxi    # Oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_exactbase"
os.makedirs(OUT, exist_ok=True)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""
SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:compatSetting w:name="compatibilityMode"
 w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="%(lat)s" w:hAnsi="%(lat)s" w:eastAsia="%(ea)s" w:cs="%(lat)s"/>
<w:sz w:val="%(sz)d"/><w:szCs w:val="%(sz)d"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

# (tag, eastAsia face, latin face, text)
FACES = [
    ("min", "ＭＳ 明朝", "Century", "本項に定める事項については関係法令及び本規程"),
    ("goth", "ＭＳ ゴシック", "Century", "本項に定める事項については関係法令及び本規程"),
    ("lat", "ＭＳ 明朝", "Times New Roman", "Handgloves quick brown fox jumps over lazy"),
]
SIZES = [16, 21, 28]          # half-points: 8pt, 10.5pt, 14pt
LINES = [120, 160, 200, 240, 320, 400]   # twentieths: 6, 8, 10, 12, 16, 20pt

# ★S1261 was measured on grid-LESS pages and then regressed real documents that
# all carry a typed docGrid (34140b: `<w:docGrid w:type="lines" linePitch="360">`,
# three exact values 280/300/340). Sweep the grid as a third axis: if Word's
# baseline stops being 0.8 x line once a grid is present, that is the scope
# condition the first sweep could not see.
GRIDS = {"nogrid": None, "grid360": 360}

ARMS = [("%s_%s_sz%d_l%d" % (g, f[0], sz, ln), f, sz, ln, GRIDS[g])
        for g in GRIDS for f in FACES for sz in SIZES for ln in LINES]


def build(tag, face, sz, line, grid=None):
    ea, lat, text = face[1], face[2], face[3]
    body = ""
    for i in range(6):
        body += ('<w:p><w:pPr><w:spacing w:after="0" w:line="%d" w:lineRule="exact"/></w:pPr>'
                 '<w:r><w:t xml:space="preserve">L%d %s</w:t></w:r></w:p>' % (line, i, text))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           "<w:body>\n" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"'
           ' w:header="851" w:footer="992" w:gutter="0"/>'
           + ('<w:docGrid w:type="lines" w:linePitch="%d"/>' % grid if grid else '')
           + "</w:sectPr></w:body></w:document>")
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES % {"lat": lat, "ea": ea, "sz": sz})
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    for a in ARMS:
        build(*a)
    print("built %d arms in %s" % (len(ARMS), OUT))
