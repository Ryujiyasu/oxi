# -*- coding: utf-8 -*-
"""Which structure lets Word roll a note past its reference's page?

`_pb_fnroom` (32 arms, single-line ref paragraphs, widowControl default ON)
found Word rolling in ZERO arms -- it always moves the reference line instead.
Yet S900 was derived from 81e80 + `_pb_fnarea`, where a roll WAS measured. The
two differ in exactly two ways, so sweep both:

    seg    0 = the refs sit on a one-line paragraph
           2 = the refs sit on the LAST line of a three-line paragraph
    widow  1 = widowControl default ON      0 = <w:widowControl w:val="0"/>

If Word rolls only when seg=2, the discriminator is "the line is a later line
of a paragraph", not S900's `prior_fill` (area already filled).

    python tools/metrics/_pb_fnroom2_gen.py
    python tools/metrics/_pb_fnroom2_read.py word
    python tools/metrics/_pb_fnroom2_read.py oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_fndefer_gen import CT, RELS, DRELS, SETTINGS, SEP

OUT = r"C:\tmp\pb_fnroom2"
os.makedirs(OUT, exist_ok=True)

NPRIOR = 2
NOWN = [1, 2]
# seg=2 costs two extra body lines, so shift its filler down to keep the
# reference line sweeping the same band of the page.
NFILL = {0: [45, 46, 47, 48], 2: [43, 44, 45, 46]}
WIDOW = [1, 0]
NOTE = ("Note %d: a single line of footnote text, long enough to fill one line "
        "of the note area but not two. ")
WORD_ = "wwwwwwww "

ARMS = [("f%d_s%d_w%d_o%d" % (f, s, w, o), f, s, w, o)
        for s in (0, 2) for f in NFILL[s] for w in WIDOW for o in NOWN]


def styles(widow):
    wc = "" if widow else '<w:widowControl w:val="0"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
            '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr>' + wc + '</w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/><w:pPr>' + wc + '</w:pPr></w:style>'
            '</w:styles>')


def build(tag, nfill, seg, widow, nown):
    body = ""
    for i in range(nfill):
        body += ('<w:p><w:r><w:t xml:space="preserve">FILL%03d line</w:t></w:r></w:p>'
                 % (i + 1))
    nid, ids = 2, []
    for k in range(NPRIOR):
        body += ('<w:p><w:r><w:t xml:space="preserve">PRIOR%02d line</w:t></w:r>'
                 '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                 '<w:footnoteReference w:id="%d"/></w:r></w:p>' % (k + 1, nid))
        ids.append(nid)
        nid += 1
    lead = "".join("P%02d %s" % (k + 1, WORD_ * 8) for k in range(seg))
    runs = ('<w:r><w:t xml:space="preserve">%sFINAL line</w:t></w:r>' % lead)
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
        z.writestr("word/styles.xml", styles(widow))
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/footnotes.xml", footnotes)
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    for a in ARMS:
        build(*a)
    print("built %d arms in %s" % (len(ARMS), OUT))
