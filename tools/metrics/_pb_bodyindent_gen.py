# -*- coding: utf-8 -*-
"""Where does Word's blank line between these paragraphs come from?

`legal__001a2c7f07cd358f` sets NO spacing anywhere -- docDefaults `pPrDefault`
is empty, `Normal` has only `w:ind`/`w:jc`, `BodyTextIndent` only `w:ind`, and
the numbering level only `w:ind` -- yet Word's own PDF advances 23.04pt between
consecutive single-line paragraphs whose line pitch is 11.52. Exactly one blank
line, from nothing that is declared.

The document's distinguishing marks, swept here one at a time:

  grid    <w:docGrid w:linePitch="360"/> with no w:type
  snap    <w:snapToGrid w:val="0"/> in Normal's rPr (a RUN property)
  num     the paragraphs are numbered (ilvl 2) with a direct w:ind override
  jc      Normal sets w:jc="both"

    python tools/metrics/_pb_bodyindent_gen.py
    python tools/metrics/_pb_bodyindent_read.py word|oxi
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = r"C:\tmp\pb_bodyindent"
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
NUMBERING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<w:numbering ' + W_NS + '>'
             '<w:abstractNum w:abstractNumId="0">'
             '<w:lvl w:ilvl="0"><w:start w:val="7"/><w:numFmt w:val="decimal"/>'
             '<w:lvlText w:val="%1"/><w:lvlJc w:val="left"/>'
             '<w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:lvl>'
             '<w:lvl w:ilvl="1"><w:start w:val="8"/><w:numFmt w:val="decimal"/>'
             '<w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/>'
             '<w:pPr><w:ind w:left="1080" w:hanging="720"/></w:pPr></w:lvl>'
             '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
             '<w:lvlText w:val="%1.%2.%3"/><w:lvlJc w:val="left"/>'
             '<w:pPr><w:ind w:left="1646" w:hanging="720"/></w:pPr></w:lvl>'
             '</w:abstractNum>'
             '<w:num w:numId="12"><w:abstractNumId w:val="0"/></w:num>'
             '</w:numbering>')

GRIDS = [360]
SNAPS = [1]
NUMS = [1]
JCS = [1]
# The real document's docDefaults has NO <w:pPrDefault> element at all, where
# the first cut of this probe wrote an empty one. A faithful slice of the real
# file DOES reproduce the inserted blank line, so the difference has to be in a
# part the hand-written arms got wrong -- this is the first candidate.
PPRDEFS = [0, 1]
ARMS = [("g%d_s%d_n%d_j%d_d%d" % (g, s, n, j, d), g, s, n, j, d)
        for g in GRIDS for s in SNAPS for n in NUMS for j in JCS for d in PPRDEFS]
TEXT = ("All staff must undergo risk stratification training before using "
        "risk stratification reports.")


def styles(snap, jc, pprdef):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles ' + W_NS + '>'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '</w:rPr></w:rPrDefault>'
            + ('<w:pPrDefault/>' if pprdef else '') +
            '</w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/><w:pPr><w:ind w:left="567"/>'
            + ('<w:jc w:val="both"/>' if jc else "") +
            '</w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            + ('<w:snapToGrid w:val="0"/>' if snap else "") +
            '</w:rPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="BodyTextIndent">'
            '<w:name w:val="Body Text Indent"/><w:basedOn w:val="Normal"/>'
            '<w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:style>'
            '</w:styles>')


def build(tag, grid, snap, num, jc, pprdef):
    npr = ('<w:numPr><w:ilvl w:val="2"/><w:numId w:val="12"/></w:numPr>'
           '<w:ind w:left="1276"/>' if num else "")
    body = ""
    for k in range(3):
        body += ('<w:p><w:pPr><w:pStyle w:val="BodyTextIndent"/>' + npr +
                 '</w:pPr><w:r><w:t xml:space="preserve">P%d %s</w:t></w:r></w:p>'
                 % (k + 1, TEXT))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document ' + W_NS + '><w:body>' + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
           '<w:pgMar w:top="2234" w:right="1440" w:bottom="1440" w:left="1276" '
           'w:header="993" w:footer="709" w:gutter="0"/>'
           + ('<w:docGrid w:linePitch="%d"/>' % grid if grid else "") +
           '</w:sectPr></w:body></w:document>')
    path = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles(snap, jc, pprdef))
        z.writestr("word/numbering.xml", NUMBERING)
        z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    for a in ARMS:
        build(*a)
    print("built %d arms in %s" % (len(ARMS), OUT))
