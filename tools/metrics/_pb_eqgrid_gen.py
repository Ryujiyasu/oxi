# -*- coding: utf-8 -*-
"""Does a DISPLAY equation snap to whole grid lines, or is it always 3?

`probeomml_equations` measures Word giving EVERY one of its 7 display equations
exactly 3.000 grid cells (54.00pt on an 18.00pt grid) where Oxi spends the raw
2.865 (51.57). One document cannot tell a snap from a constant, so sweep the
equation's natural HEIGHT against a fixed grid:

    plain   x=1                     one line tall
    frac    a/b                     numerator + denominator
    nary    SIGMA with under/over limits
    both    the witness's shape (fraction = SIGMA with limits)
    stack   a fraction whose numerator is itself a fraction
    deep    three fractions deep

Ceil-to-grid predicts 1, 2, 3, 3, 3, 4 cells; a constant predicts 3 for all.
The grid is also swept (linePitch 360 = 18pt, 240 = 12pt, 480 = 24pt) so the
answer cannot be an artifact of one cell size.

    python tools/metrics/_pb_eqgrid_gen.py         # build
    python tools/metrics/_pb_eqgrid_read.py word   # Word truth (COM -> PDF)
    python tools/metrics/_pb_eqgrid_read.py oxi    # Oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_eqgrid"
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
<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝" w:cs="Century"/>
<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

MRPR = ('<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>'
        '<w:sz w:val="21"/></w:rPr>')
CTRL = "<m:ctrlPr>%s</m:ctrlPr>" % MRPR


def mr(t):
    return "<m:r>%s<m:t>%s</m:t></m:r>" % (MRPR, t)


def frac(num, den):
    return ("<m:f><m:fPr>%s</m:fPr><m:num>%s</m:num><m:den>%s</m:den></m:f>"
            % (CTRL, num, den))


NARY = ('<m:nary><m:naryPr><m:chr m:val="\u2211"/><m:limLoc m:val="undOvr"/>%s'
        "</m:naryPr><m:sub>%s</m:sub><m:sup>%s</m:sup><m:e>%s</m:e></m:nary>"
        % (CTRL, mr("i=1"), mr("n"), mr("x")))

SHAPES = {
    "plain": mr("x=1"),
    "frac":  frac(mr("a"), mr("b")),
    "nary":  NARY,
    "both":  frac(mr("a+b"), mr("c-2")) + mr("=") + NARY,
    "stack": frac(frac(mr("a"), mr("b")), mr("c")),
    "deep":  frac(frac(frac(mr("a"), mr("b")), mr("c")), mr("d")),
}
GRIDS = {"g360": 360, "g240": 240, "g480": 480}

BODY = ("<w:p><w:r><w:t>%s</w:t></w:r></w:p>")
LINE = "本項に定める事項については関係法令及び本規程の趣旨に照らし処理する。"

ARMS = [("%s_%s" % (g, s), GRIDS[g], SHAPES[s]) for g in GRIDS for s in SHAPES]


def build(tag, pitch, math):
    body = ""
    for i in range(3):
        body += BODY % ("BEFORE%d %s" % (i, LINE))
    body += ('<w:p><m:oMathPara><m:oMath>%s</m:oMath></m:oMathPara></w:p>' % math)
    for i in range(3):
        body += BODY % ("AFTER%d %s" % (i, LINE))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
           ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
           "<w:body>\n" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1418" w:right="1134" w:bottom="1134" w:left="1134"'
           ' w:header="720" w:footer="720" w:gutter="0"/>'
           '<w:docGrid w:type="lines" w:linePitch="%d"/>'
           "</w:sectPr></w:body></w:document>" % pitch)
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    for a in ARMS:
        build(*a)
    print("built %d arms in %s" % (len(ARMS), OUT))
