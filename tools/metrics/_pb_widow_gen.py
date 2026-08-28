# -*- coding: utf-8 -*-
"""Widow/orphan control at a page break: body paragraph vs table cell.

The R33 v2 control (`_pb_kntbl2_*`) shows Word splitting a 5-line table cell
3/2 where Oxi splits 4/1, and moving the whole cell where Oxi leaves 1/4 --
exactly `w:widowControl` (default ON: never leave one line behind, never carry
one line over).  This probe asks whether Oxi honours it anywhere, by running the
same sweep on a BODY paragraph and on a CELL paragraph, with the flag left
default and explicitly turned off.

  BODY / CELL      widowControl not mentioned (Word default = ON)
  BODYOFF / CELLOFF  <w:widowControl w:val="0"> on the split paragraph
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_widow"
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

# Five lines in a 109.2pt column (2400tw grid, 108tw margins), Arial 10pt --
# the same string the R33 probe wraps, so both shapes split the same content.
TALL = ("Non contentious or common form probate business and other "
        "proceedings of a like kind")

def para(off, width_pt=None):
    """The paragraph under test. width via ind when used outside a table."""
    wc = '<w:widowControl w:val="0"/>' if off else ""
    ind = (f'<w:ind w:right="{int(round((481.9 - width_pt) * 20))}"/>'
           if width_pt else "")
    return (f'<w:p><w:pPr>{wc}{ind}</w:pPr>'
            f'<w:r><w:t xml:space="preserve">{TALL}</w:t></w:r></w:p>')

def body_block(off):
    # squeeze the body paragraph to the same 109.2pt measure as the cell
    return para(off, 109.2)

def cell_block(off):
    tc = ('<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
          + para(off) + '</w:tc>')
    tc2 = ('<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
           '<w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>')
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>'
            '</w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>'
            '<w:tr>' + tc + tc2 + '</w:tr></w:tbl>')

SHAPES = {"BODY": (body_block, False), "BODYOFF": (body_block, True),
          "CELL": (cell_block, False), "CELLOFF": (cell_block, True)}
FILLS = list(range(52, 64))

def doc_xml(tag, nfill, block, off):
    fill = "".join(f'<w:p><w:r><w:t>F{i:03d}</w:t></w:r></w:p>'
                   for i in range(1, nfill + 1))
    return (f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>#{tag}#</w:t></w:r></w:p>
{fill}{block(off)}<w:p><w:r><w:t>AFTER</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"
 w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>""")

ARMS = [(f"{s}{n}", n, blk, off) for s, (blk, off) in SHAPES.items() for n in FILLS]

def build(tag, nfill, block, off):
    p = os.path.join(OUT, f"{tag}.docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc_xml(tag, nfill, block, off))
    return p

if __name__ == "__main__":
    for a in ARMS: build(*a)
    print(f"built {len(ARMS)} arms in {OUT}")
