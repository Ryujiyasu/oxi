# -*- coding: utf-8 -*-
"""Footnotes AND a multi-line paragraph at the same page boundary.

Three probes established that each mechanism is right on its own:

    _pb_fndefer   9 arms   Oxi = Word (Word never stops the body at a reference)
    _pb_fncarry   9 arms   Oxi = Word (Word pushes body rather than carry a note)
    _pb_widow    30 arms   Oxi = Word (widowControl turns "leave 1" into "leave 0")

Yet `reports__0018715b4769984f` p5 still splits differently: Word ends its body
63.96pt above the separator where its own normal gap is 12.7-17.6, and Oxi ends
11.44 above it. The page that does this carries three notes, has the fourth
reference on its last body line, and is followed by a 3-line paragraph -- i.e.
both mechanisms act on the same break.

So compose them: grow the note area with `nrefs` one-line notes, then place a
paragraph of P lines at the boundary. The note area's height decides how many of
P fit, and widowControl decides what happens to the remainder.

    python tools/metrics/_pb_fnwidow_gen.py         # build
    python tools/metrics/_pb_fnwidow_read.py word   # Word truth (COM -> PDF)
    python tools/metrics/_pb_fnwidow_read.py oxi    # Oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_fnwidow"
os.makedirs(OUT, exist_ok=True)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fndefer_gen import CT, RELS, DRELS, SETTINGS, STYLES, SEP

NFILL = [40, 42, 44, 46]
NREFS = [1, 3]
PLINES = [2, 3]
NOTE = ("Note %d: a single line of footnote text, long enough to fill one line "
        "of the note area but not two. ")
WORD_ = "wwwwwwww "

ARMS = [("f%d_r%d_p%d" % (f, r, p), f, r, p)
        for f in NFILL for r in NREFS for p in PLINES]


def build(tag, nfill, nrefs, plines):
    body = ""
    for i in range(nfill):
        body += ('<w:p><w:r><w:t xml:space="preserve">FILL%03d line</w:t></w:r></w:p>'
                 % (i + 1))
    for k in range(nrefs):
        body += ('<w:p><w:r><w:t xml:space="preserve">REF%02d line with a note</w:t></w:r>'
                 '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                 '<w:footnoteReference w:id="%d"/></w:r></w:p>' % (k + 1, k + 2))
    text = "".join("P%02d %s" % (k + 1, WORD_ * 8) for k in range(plines))
    body += ('<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text)
    for i in range(8):
        body += ('<w:p><w:r><w:t xml:space="preserve">TAIL%03d line</w:t></w:r></w:p>'
                 % (i + 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           "<w:body>\n" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
           ' w:header="708" w:footer="708" w:gutter="0"/>'
           "</w:sectPr></w:body></w:document>")
    notes = ""
    for k in range(nrefs):
        notes += ('<w:footnote w:id="%d"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
                  'w:lineRule="auto"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>'
                  '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
                  '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:footnote>'
                  % (k + 2, NOTE % (k + 1)))
    footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                 + SEP + notes + "</w:footnotes>")
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/footnotes.xml", footnotes)
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    for a in ARMS:
        build(*a)
    print("built %d arms in %s" % (len(ARMS), OUT))
