# -*- coding: utf-8 -*-
"""What cell margin does Word use when the document declares none? (R33 control)

R33's BARE control disagreed with Word by one filler line and the archive filed
it as "a plain table's page-split capacity is off by one row".  The y-diagnosis
(_pb_kntbl_ydiag.py) says otherwise: every line tracks Word within 0.6pt until
the tall cell, where Word sets 4 lines and Oxi 5.  Word fits ink of 115.59pt in
a 120pt column; Oxi's fallback pad is 4.95pt a side (99tw), leaving 110.1.

The probe's styles.xml has no `TableNormal` style -- so this may be the minimal-
docx degradation trap, not an Oxi defect.  Sweep the declaration:

  S0  no TableNormal, no tblCellMar          <- what the R33 probe is
  S1  TableNormal w/ tblCellMar 108          <- what real Word documents carry
  S2  no TableNormal, tblPr tblCellMar 108
  S3  no TableNormal, tblPr tblCellMar 0
  S4  no TableNormal, tblPr tblCellMar 12    <- the value S0 appears to use

Read the first cell's text ORIGIN (= pad_l exactly, no side bearing) and how the
tall string wraps.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_cellmar"
os.makedirs(OUT, exist_ok=True)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
DRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

def tblnormal(ind, mar):
    """Default table style. ind=None omits <w:tblInd>; mar=None omits <w:tblCellMar>."""
    i = f'<w:tblInd w:w="{ind}" w:type="dxa"/>' if ind is not None else ""
    m = (f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="{mar}" w:type="dxa"/>'
         f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="{mar}" w:type="dxa"/></w:tblCellMar>'
         ) if mar is not None else ""
    return ('<w:style w:type="table" w:default="1" w:styleId="TableNormal">'
            '<w:name w:val="Normal Table"/><w:tblPr>' + i + m + '</w:tblPr></w:style>')

def styles(tn):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>
<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
{tblnormal(*tn) if tn else ''}
</w:styles>"""

TALL = ("Non contentious or common form probate business and other "
        "proceedings of a like kind")

def cellmar(tw):
    if tw is None: return ""
    return (f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="{tw}" w:type="dxa"/>'
            f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="{tw}" w:type="dxa"/></w:tblCellMar>')

def doc_xml(tag, tw, tind):
    tc = ('<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
          f'<w:p><w:r><w:t xml:space="preserve">{TALL}</w:t></w:r></w:p></w:tc>')
    tc2 = ('<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
           '<w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>')
    ti = f'<w:tblInd w:w="{tind}" w:type="dxa"/>' if tind is not None else ""
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>'
           + ti + cellmar(tw) +
           '</w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>'
           '<w:tr>' + tc + tc2 + '</w:tr></w:tbl>')
    return (f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>#{tag}#</w:t></w:r></w:p>
{tbl}<w:p><w:r><w:t>AFTER</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"
 w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>""")

#  tag : (default-table-style (tblInd, tblCellMar) or None, tblPr tblCellMar, tblPr tblInd)
ARMS = {
    "S0": (None,        None, None),   # nothing declared anywhere
    "S1": ((0, 108),    None, None),   # real Word's TableNormal
    "S2": (None,        108,  None),   # margin from tblPr only
    "S3": (None,        0,    None),
    "S4": (None,        12,   None),
    "S5": ((0, 108),    0,    None),
    "S6": ((None, 108), None, None),   # style margin, NO tblInd  -> isolates tblInd
    "S7": (None,        108,  0),      # tblPr margin + tblPr tblInd=0
    "S8": ((0, None),   None, None),   # style tblInd only, no margin anywhere
    "S9": ((0, 108),    108,  None),   # style + same margin restated on the table
}

def build(tag):
    tn, tw, tind = ARMS[tag]
    p = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles(tn))
        z.writestr("word/document.xml", doc_xml(tag, tw, tind))
    return p

if __name__ == "__main__":
    for t in ARMS: build(t)
    print(f"built {len(ARMS)} arms in {OUT}")
