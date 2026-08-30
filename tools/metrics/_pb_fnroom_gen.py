# -*- coding: utf-8 -*-
"""Does a reference line stay when NONE of its own notes can start on the page?

S900 lets a line keep its place and roll its notes to the next page when
EITHER at least one of the line's own notes places, OR earlier notes already
fill the area ("moving the line would free nothing"). The composite probe
`_pb_fnwidow` f46_r3_p* showed Word moving such a line anyway, so the second
disjunct is under test here.

Arms: `nprior` one-note reference lines fill the note area first, then ONE
final line carries `nown` references. Sweeping `nfill` walks the boundary so
that 2, 1 and 0 of the final line's own notes fit.

    python tools/metrics/_pb_fnroom_gen.py
    python tools/metrics/_pb_fnroom_read.py word
    python tools/metrics/_pb_fnroom_read.py oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fndefer_gen import CT, RELS, DRELS, SETTINGS, STYLES, SEP

OUT = r"C:\tmp\pb_fnroom"
os.makedirs(OUT, exist_ok=True)

NFILL = [40, 42, 43, 44, 45, 46, 47, 48]
NPRIOR = [0, 2]
NOWN = [1, 2]
NOTE = ("Note %d: a single line of footnote text, long enough to fill one line "
        "of the note area but not two. ")

ARMS = [("f%d_q%d_o%d" % (f, q, o), f, q, o)
        for f in NFILL for q in NPRIOR for o in NOWN]


def build(tag, nfill, nprior, nown):
    body = ""
    for i in range(nfill):
        body += ('<w:p><w:r><w:t xml:space="preserve">FILL%03d line</w:t></w:r></w:p>'
                 % (i + 1))
    nid = 2
    ids = []
    for k in range(nprior):
        body += ('<w:p><w:r><w:t xml:space="preserve">PRIOR%02d line</w:t></w:r>'
                 '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                 '<w:footnoteReference w:id="%d"/></w:r></w:p>' % (k + 1, nid))
        ids.append(nid)
        nid += 1
    # the line under test: one paragraph carrying `nown` references
    runs = '<w:r><w:t xml:space="preserve">FINAL line</w:t></w:r>'
    for k in range(nown):
        runs += ('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                 '<w:footnoteReference w:id="%d"/></w:r>' % nid)
        ids.append(nid)
        nid += 1
    body += "<w:p>" + runs + "</w:p>"
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
    for n, fid in enumerate(ids):
        notes += ('<w:footnote w:id="%d"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
                  'w:lineRule="auto"/><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>'
                  '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
                  '<w:t xml:space="preserve">%s</w:t></w:r></w:p></w:footnote>'
                  % (fid, NOTE % (n + 1)))
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
