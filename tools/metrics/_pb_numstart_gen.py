# -*- coding: utf-8 -*-
"""What does an UNUSED parent numbering level contribute to a child's marker?

`legal__001a2c7f07cd358f` numbers two paragraphs at ilvl 2 of a list whose
ilvl 0 declares `w:start="7"` and ilvl 1 `w:start="8"`; no paragraph in the
document ever uses those two levels. Oxi renders the markers as `0.0.1` and
`0.0.2`. A level that has never been incremented should still read as its
declared start, so the expected marker is `7.8.1`.

Arms: the parent starts, whether a parent level is actually used, and the
child's own start -- so the answer separates "unused parent = start" from
"unused parent = 0" and from "unused parent = 1".

    python tools/metrics/_pb_numstart_gen.py
    python tools/metrics/_pb_numstart_read.py word|oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = r"C:\tmp\pb_numstart"
W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
         '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles ' + W_NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
          '<w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/></w:style></w:styles>')

# (parent0 start, parent1 start, child start, which levels the body uses)
ARMS = [
    ("s7_s8_used0", 7, 8, 1, [2]),
    ("s7_s8_used01", 7, 8, 1, [0, 1, 2]),
    ("s1_s1_used0", 1, 1, 1, [2]),
    ("s3_s5_used0", 3, 5, 4, [2]),
    ("s7_s8_used1", 7, 8, 1, [1, 2]),
]


def numbering(p0, p1, ch):
    lv = ""
    for i, (st, txt) in enumerate(((p0, "%1"), (p1, "%1.%2"), (ch, "%1.%2.%3"))):
        lv += ('<w:lvl w:ilvl="%d"><w:start w:val="%d"/><w:numFmt w:val="decimal"/>'
               '<w:lvlText w:val="%s"/><w:lvlJc w:val="left"/>'
               '<w:pPr><w:ind w:left="%d" w:hanging="720"/></w:pPr></w:lvl>'
               % (i, st, txt, 720 * (i + 1)))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:numbering ' + W_NS + '>'
            '<w:abstractNum w:abstractNumId="0">' + lv + '</w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
            '</w:numbering>')


def build(tag, p0, p1, ch, used):
    body = ""
    for lvl in used:
        for k in range(2):
            body += ('<w:p><w:pPr><w:numPr><w:ilvl w:val="%d"/><w:numId w:val="1"/>'
                     '</w:numPr></w:pPr><w:r><w:t xml:space="preserve">L%d item %d</w:t>'
                     '</w:r></w:p>' % (lvl, lvl, k + 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document ' + W_NS + '><w:body>' + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
           "</w:body></w:document>")
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", numbering(p0, p1, ch))
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    for a in ARMS:
        build(*a)
    print("built %d arms in %s" % (len(ARMS), OUT))
