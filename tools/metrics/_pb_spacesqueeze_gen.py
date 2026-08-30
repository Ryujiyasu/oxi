# -*- coding: utf-8 -*-
"""How far will Word SQUEEZE inter-word spaces to avoid a wrap?

Measured on `creative__009790431a821d2f` (Calibri 11.04): the natural space is
2.430pt, justified middle lines STRETCH it to 3.0-4.5, and the paragraph's last
line -- which is not justified at all -- is set with the space at **1.875pt**,
77.2% of natural. So Word squeezes to keep a word on the line. That number is
the amount that line HAPPENED to need, not the limit.

To find the limit, do not try to control the overflow finely (a proportional
font makes that awkward). Instead sweep the filler length one character at a
time and find the LAST arm that still fits on one line: the space advance in
that arm is as far as Word was willing to go. Sweeping the number of spaces as
well separates a per-space floor ("each space may shrink to X% of natural")
from a per-line budget ("the line may borrow at most Y points in total").

    python tools/metrics/_pb_spacesqueeze_gen.py         # build
    python tools/metrics/_pb_spacesqueeze_read.py word   # Word truth (COM -> PDF)
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_spacesqueeze"
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
<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
<w:sz w:val="%(sz)d"/><w:szCs w:val="%(sz)d"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

# `n` words of `w` letters, then a tail word of `L` letters. Sweeping L one
# letter at a time steps the demand by ~5pt (Calibri 'n' at 11pt), and the
# boundary arm is whichever L last fits.
# Calibri 11: 'n' = 5.760pt, space = 2.514pt (both read out of Word's own PDF).
# Pick (word length, word count) so the tail word lands ON the boundary of the
# 451.32pt column, and give the three shapes very different SPACE COUNTS -- that
# is what separates "each space may shrink to X% of natural" from "the line may
# borrow at most Y points in total".
#   long  : 5 words x 14 chars ->  5 spaces, filled to ~415.8 before the tail
#   mid   : 9 words x  8 chars ->  9 spaces, ~437.4
#   short : 22 words x 3 chars -> 22 spaces, ~435.4
SHAPES = {"long": (14, 5), "mid": (8, 9), "short": (3, 22)}
LS = list(range(1, 13))
SIZES = {"sz22": 22}
JC = {"both": '<w:jc w:val="both"/>', "left": ""}

ARMS = [("%s_%s_%s_L%d" % (s, j, w, L), SIZES[s], JC[j], SHAPES[w], L)
        for s in SIZES for j in JC for w in SHAPES for L in LS]


def build(tag, sz, jc, shape, L):
    wlen, nwords = shape
    body = " ".join(["n" * wlen] * nwords) + " " + ("n" * L)
    p = ('<w:p><w:pPr>%s</w:pPr><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
         % (jc, body))
    p += "<w:p><w:r><w:t>AFTER</w:t></w:r></w:p>"
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           "<w:body>\n" + p +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
           ' w:header="708" w:footer="708" w:gutter="0"/>'
           "</w:sectPr></w:body></w:document>")
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES % {"sz": sz})
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    for a in ARMS:
        build(*a)
    print("built %d arms in %s  (column 451.32pt)" % (len(ARMS), OUT))
