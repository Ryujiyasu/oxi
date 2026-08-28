# -*- coding: utf-8 -*-
"""keepNext-chain table probe, v2 — FAITHFUL package (R33 follow-up).

v1 (`_pb_kntbl_gen.py`) shipped a styles.xml with no default table style and no
settings.xml.  Its BARE control disagreed with Word by one filler line and the
archive filed that as "a plain table's page-split capacity is off by one row".
It is not: `_pb_cellmar_read.py` measures Word using ~0 cell margin when nothing
declares one, against Oxi's 99tw fallback, so the tall cell wrapped to 5 lines in
Oxi and 4 in Word.  A one-line taller table, not a page-capacity error.
(= [[probe_minimal_docx_degraded]] again.)

v2 adds what every real Word document carries and v1 omitted:
  * a default table style `TableNormal` (tblInd 0, tblCellMar 108 L/R)
  * settings.xml with compatibilityMode 15
Both matter: the margin sets the wrap budget, and the compat mode decides whether
the leading cell absorbs its own margin (S496).

Shapes and sweep are unchanged from v1 so the two are directly comparable.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_kntbl2"
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

KN = "<w:keepNext/>"
TALL = ("Non contentious or common form probate business and other "
        "proceedings of a like kind")

def cell(text, kn):
    pr = f"<w:pPr>{KN if kn else ''}</w:pPr>"
    return ('<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
            f'<w:p>{pr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>')

def table(row_kn):
    rows = [("HDR-A", "HDR-B", row_kn), ("R1-A", "R1-B", row_kn),
            (TALL, "R2-B", row_kn), ("R3-A", "R3-B", False),
            ("R4-A", "R4-B", False)]
    trs = "".join("<w:tr>" + cell(a, kn) + cell(b, kn) + "</w:tr>"
                  for a, b, kn in rows)
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>'
            + trs + "</w:tbl>")

def doc_xml(nfill, cap_kn, row_kn, tag):
    fill = "".join(f'<w:p><w:r><w:t>F{i:03d}</w:t></w:r></w:p>'
                   for i in range(1, nfill + 1))
    cap = (f'<w:p><w:pPr>{KN if cap_kn else ""}<w:jc w:val="center"/></w:pPr>'
           f'<w:r><w:t>CAPTION</w:t></w:r></w:p>')
    return (f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>#{tag}#</w:t></w:r></w:p>
{fill}{cap}{table(row_kn)}
<w:p><w:r><w:t>AFTER</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"
 w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>""")

SHAPES = {"KN": (True, True), "NOKN": (True, False), "BARE": (False, False)}
FILLS = list(range(46, 67))

def build(tag, nfill, cap_kn, row_kn):
    p = os.path.join(OUT, f"{tag}.docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc_xml(nfill, cap_kn, row_kn, tag))
    return p

ARMS = [(f"{s}{n}", n, ck, rk) for s, (ck, rk) in SHAPES.items() for n in FILLS]

if __name__ == "__main__":
    for tag, n, ck, rk in ARMS:
        build(tag, n, ck, rk)
    print(f"built {len(ARMS)} arms in {OUT}")
