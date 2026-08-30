# -*- coding: utf-8 -*-
"""Which paragraph loses its `w:after`, and why?

`legal__000ad039fbd3f8b6` diverges from Word in ONE step of 35.70pt on page 14:
the heading "Authors Cited" carries `w:spacing w:after="720"` (36pt) and
`w:keepNext`, and Word leaves 49.50pt to the next paragraph (line + 36) where
Oxi leaves 13.80 (line only). Four other documents fail the same way and in the
same direction, so isolate the trigger rather than patch the document.

Axes on the first paragraph: keepNext on/off, the after value, and whether the
next paragraph shares its style (contextualSpacing is not declared anywhere in
that document, but the pair does share a custom style, so rule it out).

    python tools/metrics/_pb_afterkeep_gen.py
    python tools/metrics/_pb_afterkeep_read.py word|oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = r"C:\tmp\pb_afterkeep"
W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '</Relationships>')
# The real document's shape: Normal sz 24, a custom style based on it that sets
# double spacing, and paragraphs that override the line rule back to single.
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles ' + W_NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Dbl">'
          '<w:name w:val="SCC.Normal.DoubleSpacing"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:spacing w:line="480" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>'
          '</w:style></w:styles>')

KEEPS = [0, 1]
AFTERS = [240, 720]
SAME_STYLE = [0, 1]
# In the real document the heading that loses its after-space is the FIRST
# paragraph on its page, so sweep that too: `pageBreakBefore` puts HEAD at a
# page top without changing anything else about it.
TOPS = [0, 1]
# FILLS puts HEAD near the page bottom: at 43-45 filler lines HEAD still fits
# but NEXT does not, so a keepNext pair has to MOVE together. That is the shape
# the real document is in, and the one the flat arms never reach.
# 39-41 is the band where the keepNext pair check passes (it omits the
# collapsed gap when the follower carries direct spacing) but the real layout
# overflows -- the S960 back-pull path, which 43-45 never reaches because the
# look-ahead moves the pair early there.
FILLS = [39, 40, 41, 43, 45]
# In the real document HEAD is itself preceded by an after=720 paragraph, so it
# renders 36pt below its own start_y. PRES gives the last filler that same
# after, which is what stops the keepNext look-ahead from moving HEAD early.
PRES = [0, 720]
ARMS = [("k%d_a%d_s%d_t%d_f%d_p%d" % (k, a, s, t, f, pr), k, a, s, t, f, pr)
        for k in KEEPS for a in AFTERS for s in SAME_STYLE for t in TOPS
        for f in FILLS for pr in PRES]


def build(tag, keepnext, after_tw, same_style, at_top, nfill, pre_tw):
    st = '<w:pStyle w:val="Dbl"/>'
    fill = "".join(
        '<w:p><w:pPr><w:spacing w:after="%d" w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">FILL%03d line</w:t></w:r></w:p>'
        % (pre_tw if i == nfill - 1 else 0, i + 1)
        for i in range(nfill))
    first = ('<w:p><w:pPr>' + st +
             ("<w:pageBreakBefore/>" if at_top else "") +
             ("<w:keepNext/>" if keepnext else "") +
             '<w:spacing w:after="%d" w:line="240" w:lineRule="auto"/>'
             '<w:rPr><w:b/></w:rPr></w:pPr>'
             '<w:r><w:rPr><w:b/></w:rPr>'
             '<w:t xml:space="preserve">HEAD line</w:t></w:r></w:p>' % after_tw)
    second = ('<w:p><w:pPr>' + (st if same_style else "") +
              '<w:spacing w:after="240" w:line="240" w:lineRule="auto"/></w:pPr>'
              '<w:r><w:t xml:space="preserve">NEXT line</w:t></w:r></w:p>')
    third = ('<w:p><w:pPr><w:spacing w:after="240" w:line="240" w:lineRule="auto"/></w:pPr>'
             '<w:r><w:t xml:space="preserve">TAIL line</w:t></w:r></w:p>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document ' + W_NS + '><w:body>' + fill + first + second + third +
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
           "</w:body></w:document>")
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    for a in ARMS:
        build(*a)
    print("built %d arms in %s" % (len(ARMS), OUT))
