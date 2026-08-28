# -*- coding: utf-8 -*-
"""A deleted paragraph MARK joins two paragraphs -- whose properties survive?

`<w:pPr><w:rPr><w:del/></w:rPr></w:pPr>` deletes the pilcrow. Accepting that
revision merges the paragraph with the one after it, so the break and its
paragraph spacing stop existing. `legal__0010437a7f75f636` carries 61 of them
and Word's accepted truth has exactly 62 fewer paragraphs than its as-authored
one -- but that document cannot say WHICH paragraph's properties the merged
paragraph keeps, because the pairs share a style.

Give the two paragraphs different left indents and read the merged line's x:

  CTRL    no deleted mark                 -- two paragraphs, two indents
  MERGE   head mark deleted               -- one paragraph; x says whose pPr won
  CHAIN   two consecutive marks deleted   -- three paragraphs collapse to one
  LAST    the body's final mark deleted   -- nothing to join; must not vanish
  CELL    the mark of a cell's last paragraph deleted -- same, inside a table

Read with `_pb_delmark_read.py word|oxi`.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
OUT = r"C:\tmp\pb_delmark"
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

DEL = ('<w:rPr><w:del w:id="%d" w:author="Probe" '
       'w:date="2026-08-28T00:00:00Z"/></w:rPr>')


def para(text, indent, delmark=None):
    """One paragraph. `indent` in twips; `delmark` an id to delete its mark."""
    pr = '<w:pPr><w:ind w:left="%d"/>%s</w:pPr>' % (
        indent, (DEL % delmark) if delmark else "")
    return pr.join(("<w:p>", "")) + '<w:r><w:t>%s</w:t></w:r></w:p>' % text


def cell(inner):
    return ('<w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr>'
            + inner + '</w:tc>')


def body(tag):
    if tag == "CTRL":
        return para("HEAD", 1440) + para("TAIL", 2880) + para("AFTER", 0)
    if tag == "MERGE":
        return para("HEAD", 1440, 101) + para("TAIL", 2880) + para("AFTER", 0)
    if tag == "CHAIN":
        return (para("ONE", 1440, 102) + para("TWO", 2160, 103)
                + para("THREE", 2880) + para("AFTER", 0))
    if tag == "LAST":
        return para("HEAD", 1440) + para("TAIL", 2880, 104)
    if tag == "CELL":
        inner = para("HEAD", 1440, 105) + para("TAIL", 2880, 106)
        tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
               '<w:tblLayout w:type="fixed"/></w:tblPr>'
               '<w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid>'
               '<w:tr>' + cell(inner) + '</w:tr></w:tbl>')
        return tbl + para("AFTER", 0)
    raise ValueError(tag)


TAGS = ["CTRL", "MERGE", "CHAIN", "LAST", "CELL"]


def doc_xml(tag):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body>\n<w:p><w:r><w:t>#' + tag + '#</w:t></w:r></w:p>\n'
            + body(tag) +
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"'
            ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


def build(tag):
    p = os.path.join(OUT, tag + ".docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", doc_xml(tag))
    return p


if __name__ == "__main__":
    for t in TAGS:
        build(t)
    print("built %d arms in %s" % (len(TAGS), OUT))
