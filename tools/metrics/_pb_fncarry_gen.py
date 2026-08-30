# -*- coding: utf-8 -*-
"""What happens on the page where a footnote has to be CARRIED to the next one?

`_pb_fndefer` showed Word never stops the body at a reference, and that Oxi
reproduces Word's break in all 9 arms when the page holds ONE note. The real
document that still fails (`reports__0018715b4769984f` p5) differs in exactly one
structural way: it carries three notes and defers a fourth, whose reference sits
on the page's last body line. There Oxi ends its body 11.4pt above the separator
where Word keeps 28-35.

So build the carry: several one-line notes near the foot of the page, enough that
the last one cannot fit. Each new reference grows the note area and shrinks the
body, so the arms walk into the deferral rather than being placed in it.

    python tools/metrics/_pb_fncarry_gen.py         # build
    python tools/metrics/_pb_fncarry_read.py word   # Word truth (COM -> PDF)
    python tools/metrics/_pb_fncarry_read.py oxi    # Oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_fncarry"
os.makedirs(OUT, exist_ok=True)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fndefer_gen import CT, RELS, DRELS, SETTINGS, STYLES, SEP

NBEFORE = [30, 34, 38]
NREFS = [2, 4, 6]
NOTE = ("Note %d: a single line of footnote text, long enough to fill one line "
        "of the note area but not two. ")

ARMS = [("b%d_r%d" % (b, r), b, r) for b in NBEFORE for r in NREFS]


def build(tag, nbefore, nrefs):
    body = ""
    for i in range(nbefore):
        body += ('<w:p><w:r><w:t xml:space="preserve">BODY%03d filler line</w:t>'
                 "</w:r></w:p>" % (i + 1))
    for k in range(nrefs):
        body += ('<w:p><w:r><w:t xml:space="preserve">REF%02d line with a note</w:t></w:r>'
                 '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                 '<w:footnoteReference w:id="%d"/></w:r></w:p>' % (k + 1, k + 2))
    for i in range(14):
        body += ('<w:p><w:r><w:t xml:space="preserve">AFTER%03d filler line</w:t>'
                 "</w:r></w:p>" % (i + 1))
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
