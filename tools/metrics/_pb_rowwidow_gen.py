# -*- coding: utf-8 -*-
"""Is a splitting row's widow/orphan rule per PARAGRAPH or per ROW? (S1246)

`_pb_widow_*` established that Word applies widow/orphan control when a table
row splits, and that `w:widowControl w:val="0"` turns it off.  That probe's row
held ONE tall paragraph, so "the paragraph has a lone line on one side" and "the
ROW has a lone line on one side" were the same statement.

uklocalspending p36 separates them and contradicts the per-paragraph reading:
its row has 3-line cells beside a 15-line cell, Word splits after 2 lines, and
the 3-line cells each carry a lone last line -- a paragraph widow Word does not
avoid, while the ROW keeps 2 lines above and many below.

Arms (cell A short, cell B tall, sweeping the split through the row):
  EQ     A 5 lines, B 5 lines   -- the two models agree; the earlier result
  SHORT  A 3 lines, B 9 lines   -- uklocalspending's shape, synthesised
  MED    A 5 lines, B 9 lines   -- A widows while the row still has 2 a side
A pull under SHORT/MED means the rule is per paragraph; no pull means per row.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_rowwidow"
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
<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>
<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>
<w:tblPr><w:tblInd w:w="0" w:type="dxa"/>
<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>
</w:tblPr></w:style>
</w:styles>"""


def lines_text(letter, n):
    """One word per line: 13 repeats of the letter plus a 2-digit index is
    ~95pt at Arial 10, which fits the 109.2pt measure alone but never in
    pairs -- so the paragraph has exactly n lines and each is identifiable."""
    return " ".join("%s%02d" % (letter * 13, i) for i in range(1, n + 1))


SPACING = ('<w:pPr><w:spacing w:before="120" w:after="120" w:line="0"'
           ' w:lineRule="atLeast"/></w:pPr>')


def cell(letter, n, sp=False):
    """sp=True gives the cell paragraph uklocalspending's own spacing
    (before/after 6pt, line atLeast 0) -- the largest untested difference
    between this probe and the real document that contradicts it."""
    return ('<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
            '<w:p>' + (SPACING if sp else "")
            + '<w:r><w:t xml:space="preserve">' + lines_text(letter, n)
            + '</w:t></w:r></w:p></w:tc>')


def table(na, nb, pre=False, sp=False):
    """pre=True puts a 1-line row ahead of the row under test, so the splitting
    row is an INTERIOR row rather than the table's first -- uklocalspending's
    shape, where moving the row whole is an ordinary interior move rather than
    moving the entire table off the page."""
    head = ('<w:tr>' + cell("P", 1) + cell("Q", 1) + '</w:tr>') if pre else ""
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>'
            '</w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>'
            + head + '<w:tr>' + cell("A", na, sp) + cell("B", nb, sp)
            + '</w:tr></w:tbl>')


#  tag: (A lines, B lines, preceding row, uklocalspending spacing)
SHAPES = {
    "EQ": (5, 5, False), "SHORT": (3, 9, False), "MED": (5, 9, False),
    # uklocalspending p36 splits a 2-line cell 1/1 and a 3-line cell 2/1 beside
    # a 15-line cell, with COM reporting WidowControl=-1 on every one of them.
    # TWO asks whether a 2-line cell is the exception; PRESHORT/PRETWO ask
    # whether being an interior row is.
    "TWO": (2, 9, False), "PRESHORT": (3, 9, True), "PRETWO": (2, 9, True),
    # SPSHORT/SPMED repeat SHORT/MED with that spacing: SPMED says whether the
    # SHIPPED n>=4 rule survives it, SPSHORT whether the parked n=3 region flips
    # to uklocalspending's answer.
    "SPSHORT": (3, 9, False, True), "SPMED": (5, 9, False, True),
}
FILLS = list(range(50, 64))


def doc_xml(tag, nfill, na, nb, pre, sp):
    fill = "".join('<w:p><w:r><w:t>F%03d</w:t></w:r></w:p>' % i
                   for i in range(1, nfill + 1))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body>\n<w:p><w:r><w:t>#' + tag + '#</w:t></w:r></w:p>\n'
            + fill + table(na, nb, pre, sp) + '<w:p><w:r><w:t>AFTER</w:t></w:r></w:p>\n'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>\n'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"\n'
            ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


def _shape(v):
    """(A lines, B lines, preceding row, spacing) with the tail defaulted."""
    return (list(v) + [False, False])[:4]


ARMS = [tuple(["%s%d" % (s, n), n] + _shape(v))
        for s, v in SHAPES.items() for n in FILLS]


def build(tag, nfill, na, nb, pre, sp):
    p = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc_xml(tag, nfill, na, nb, pre, sp))
    return p


if __name__ == "__main__":
    for a in ARMS:
        build(*a)
    print("built %d arms in %s" % (len(ARMS), OUT))
