# -*- coding: utf-8 -*-
"""How does Word split a paragraph across a page break? (widowControl)

Two probes now agree that Oxi's footnote-area reservation matches Word exactly
(18 arms, same last body line in every one, gaps within 0.35pt). So the residual
on `reports__0018715b4769984f` p5 is not the reservation: Word ends that page
63.96pt above the separator where its own normal gap is 12.7-17.6, and the
difference is almost exactly the 3 lines of the next paragraph, which carries no
keepNext and no keepLines. Word moved it whole.

`w:widowControl` defaults ON, and forbids leaving one line of a paragraph alone
on either side of the break. This probe measures what that actually does: fill a
page to a controlled depth, then place a paragraph of PLINES lines, and count how
many of them stay.

    python tools/metrics/_pb_widow_gen.py         # build
    python tools/metrics/_pb_widow_read.py word   # Word truth (COM -> PDF)
    python tools/metrics/_pb_widow_read.py oxi    # Oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_widow"
os.makedirs(OUT, exist_ok=True)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fndefer_gen import CT, RELS, DRELS, STYLES

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:compatSetting w:name="compatibilityMode"
 w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""

# A4 with 72pt margins holds 48 lines of Calibri 11 (14.52 each). Sweep the
# filler so that 1..5 lines of the test paragraph would fit before the break.
# A4 minus 72pt margins = 697.92pt of column; Calibri 11 at single spacing is
# 13.43pt a line, so the page takes 52. Sweep the filler across that boundary.
FILLERS = [48, 49, 50, 51, 52]
PLINES = [2, 3, 4]
WIDOW = {"on": "", "off": "<w:widowControl w:val=\"0\"/>"}
WORD_ = "wwwwwwww "

ARMS = [("%s_f%d_p%d" % (w, f, p), WIDOW[w], f, p)
        for w in WIDOW for f in FILLERS for p in PLINES]


def build(tag, widow, nfill, plines):
    body = ""
    for i in range(nfill):
        body += ('<w:p><w:pPr>%s</w:pPr><w:r><w:t xml:space="preserve">FILL%03d line</w:t>'
                 "</w:r></w:p>" % (widow, i + 1))
    # one paragraph that wraps to exactly `plines` lines: 9 words per line at
    # Calibri 11 across a 451.32pt column.
    text = "".join("P%02d %s" % (k + 1, WORD_ * 8) for k in range(plines))
    body += ('<w:p><w:pPr>%s</w:pPr><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
             % (widow, text))
    for i in range(6):
        body += ('<w:p><w:pPr>%s</w:pPr><w:r><w:t xml:space="preserve">TAIL%03d line</w:t>'
                 "</w:r></w:p>" % (widow, i + 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           "<w:body>\n" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
           ' w:header="708" w:footer="708" w:gutter="0"/>'
           "</w:sectPr></w:body></w:document>")
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
