"""Generate multiple repro variants to isolate Bug 1 cause.
Each variant changes ONE setting from baseline."""
import zipfile, os, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# Common XML pieces
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

def build_doc(extra_tblpr="", extra_ppr="", extra_rpr=""):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:type="dxa" w:w="9628"/>{extra_tblpr}</w:tblPr>
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
              {extra_ppr}
              <w:rPr>{extra_rpr}<w:rFonts w:cs="ＭＳ 明朝"/><w:color w:val="000000"/><w:spacing w:val="-1"/><w:kern w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
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

# Variants
VARIANTS = {
    "v0_baseline":     ("", "", ""),
    "v1_tcMar_540":    ('<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>', "", ""),
    "v2_no_adjustR":   ("", "", ""),  # we'll modify ppr inline
    "v3_no_autoSpace": ("", "", ""),
    "v4_default_kern": ("", "", ""),  # remove kern=0 from rPr
    "v5_no_spacing_neg": ("", "", ""),  # remove w:spacing=-1
    "v6_no_hanging":   ("", "", ""),  # remove hanging
}

# v2: build separately to remove adjustRightInd
def build_v2():
    return build_doc().replace('<w:adjustRightInd w:val="0"/>\n              ', '')

def build_v3():
    return build_doc().replace('<w:autoSpaceDE w:val="0"/>\n              <w:autoSpaceDN w:val="0"/>\n              ', '')

def build_v4():
    return build_doc().replace('<w:kern w:val="0"/>', '')

def build_v5():
    return build_doc().replace('<w:spacing w:val="-1"/>', '').replace('<w:spacing w:val="-9"/>', '')

def build_v6():
    return build_doc().replace(
        '<w:ind w:leftChars="150" w:left="5775" w:right="199" w:hangingChars="2600" w:hanging="5460"/>',
        '<w:ind w:leftChars="150" w:left="315" w:right="199"/>'
    )

builders = {
    "v0_baseline":         lambda: build_doc(),
    "v1_tcMar_540":        lambda: build_doc(extra_tblpr='<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'),
    "v2_no_adjustR":       build_v2,
    "v3_no_autoSpace":     build_v3,
    "v4_default_kern":     build_v4,
    "v5_no_spacing_neg":   build_v5,
    "v6_no_hanging":       build_v6,
}

OUT_DIR = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
for name, builder in builders.items():
    path = f"{OUT_DIR}/repro_v_{name}.docx"
    doc_xml = builder()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", REL_XML)
        z.writestr("word/_rels/document.xml.rels", DOC_REL_XML)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", STYLES_XML)
    print(f"Wrote {path}")
print(f"\nTotal: {len(builders)} variants")
