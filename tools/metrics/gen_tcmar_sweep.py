"""Sweep tblCellMar L=R values to find the threshold where Word's wrap changes."""
import zipfile, os

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
REL_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_REL_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
</w:styles>'''

def build_doc(tcmar_tw):
    """tcmar_tw: int twips for L=R, or None for no tblCellMar element."""
    if tcmar_tw is None:
        tcm = ""
    else:
        tcm = f'<w:tblCellMar><w:left w:w="{tcmar_tw}" w:type="dxa"/><w:right w:w="{tcmar_tw}" w:type="dxa"/></w:tblCellMar>'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:type="dxa" w:w="9628"/>{tcm}</w:tblPr>
      <w:tblGrid><w:gridCol w:w="9628"/></w:tblGrid>
      <w:tr><w:trHeight w:val="2000"/>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9628"/></w:tcPr>
          <w:p>
            <w:pPr>
              <w:autoSpaceDE w:val="0"/>
              <w:autoSpaceDN w:val="0"/>
              <w:adjustRightInd w:val="0"/>
              <w:snapToGrid w:val="0"/>
              <w:spacing w:line="280" w:lineRule="exact"/>
              <w:ind w:leftChars="150" w:left="5775" w:right="199" w:hangingChars="2600" w:hanging="5460"/>
              <w:rPr><w:rFonts w:cs="ＭＳ 明朝"/><w:color w:val="000000"/><w:spacing w:val="-1"/><w:kern w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r><w:rPr><w:rFonts w:cs="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:spacing w:val="-9"/><w:kern w:val="0"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">６　手数料の納付方法　</w:t></w:r>
            <w:r><w:rPr><w:rFonts w:cs="ＭＳ 明朝"/><w:color w:val="000000"/><w:spacing w:val="-9"/><w:kern w:val="0"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">　　</w:t></w:r>
            <w:r><w:rPr><w:rFonts w:cs="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:spacing w:val="-9"/><w:kern w:val="0"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">ア　</w:t></w:r>
            <w:r><w:rPr><w:rFonts w:cs="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:spacing w:val="-1"/><w:kern w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">収入印紙による納付　   イ 行政機関の長、指定独立行政法人等、独立行政法</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851"/><w:docGrid w:type="linesAndChars" w:linePitch="357"/></w:sectPr>
  </w:body>
</w:document>
'''

OUT_DIR = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
SWEEP = [None, 0, 30, 50, 80, 99, 108, 150, 200, 300, 500]
for tw in SWEEP:
    name = "none" if tw is None else f"{tw:04d}"
    path = f"{OUT_DIR}/repro_tcmar_{name}.docx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", REL_XML)
        z.writestr("word/_rels/document.xml.rels", DOC_REL_XML)
        z.writestr("word/document.xml", build_doc(tw))
        z.writestr("word/styles.xml", STYLES_XML)
    print(f"Wrote {path}")
print(f"\nTotal: {len(SWEEP)} variants")
