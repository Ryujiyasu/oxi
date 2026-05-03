"""Minimal repro: replicate 1636 item 6's exact pPr + run structure (no anchor shape).
Goal: trigger the same 1-extra-line wrap difference."""
import zipfile

DOC_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:type="dxa" w:w="9628"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="9628"/></w:tblGrid>
      <w:tr><w:trHeight w:val="2000"/>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9628"/></w:tcPr>
          <!-- Item 5 (ref) -->
          <w:p>
            <w:pPr>
              <w:wordWrap/>
              <w:snapToGrid w:val="0"/>
              <w:spacing w:beforeLines="50" w:before="136" w:line="240" w:lineRule="exact"/>
              <w:ind w:leftChars="150" w:left="315" w:right="199"/>
            </w:pPr>
            <w:r><w:t xml:space="preserve">５　手数料の額（reference, line=240/12pt）</w:t></w:r>
          </w:p>
          <!-- Item 6 (the bug target) — mimics 1636 P21 -->
          <w:p>
            <w:pPr>
              <w:autoSpaceDE w:val="0"/>
              <w:autoSpaceDN w:val="0"/>
              <w:adjustRightInd w:val="0"/>
              <w:snapToGrid w:val="0"/>
              <w:spacing w:beforeLines="50" w:before="136" w:line="280" w:lineRule="exact"/>
              <w:ind w:leftChars="150" w:left="5775" w:right="199" w:hangingChars="2600" w:hanging="5460"/>
              <w:rPr><w:rFonts w:cs="ＭＳ 明朝"/><w:color w:val="000000"/><w:spacing w:val="-1"/><w:kern w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r><w:rPr><w:rFonts w:cs="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:spacing w:val="-9"/><w:kern w:val="0"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">６　手数料の納付方法　</w:t></w:r>
            <w:r><w:rPr><w:rFonts w:cs="ＭＳ 明朝"/><w:color w:val="000000"/><w:spacing w:val="-9"/><w:kern w:val="0"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">　　</w:t></w:r>
            <w:r><w:rPr><w:rFonts w:cs="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:spacing w:val="-9"/><w:kern w:val="0"/><w:szCs w:val="21"/></w:rPr><w:t xml:space="preserve">ア　</w:t></w:r>
            <w:r><w:rPr><w:rFonts w:cs="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:spacing w:val="-1"/><w:kern w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">収入印紙による納付　   イ 行政機関の長、指定独立行政法人等、独立行政法</w:t></w:r>
          </w:p>
          <!-- Item 6 wrap continuation (P22) -->
          <w:p>
            <w:pPr>
              <w:autoSpaceDE w:val="0"/>
              <w:autoSpaceDN w:val="0"/>
              <w:adjustRightInd w:val="0"/>
              <w:snapToGrid w:val="0"/>
              <w:spacing w:line="200" w:lineRule="exact"/>
              <w:ind w:leftChars="150" w:left="5775" w:right="199"/>
              <w:rPr><w:rFonts w:cs="ＭＳ 明朝"/><w:color w:val="000000"/><w:spacing w:val="-1"/><w:kern w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
            </w:pPr>
            <w:r><w:rPr><w:rFonts w:cs="ＭＳ 明朝" w:hint="eastAsia"/><w:color w:val="000000"/><w:spacing w:val="-1"/><w:kern w:val="0"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>人統計センターがあらかじめ定めるア以外の方法</w:t></w:r>
          </w:p>
          <!-- After (item 7 placeholder) -->
          <w:p>
            <w:pPr>
              <w:spacing w:line="240" w:lineRule="exact"/>
            </w:pPr>
            <w:r><w:t>７　公表関係（reference）</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851"/><w:docGrid w:type="linesAndChars" w:linePitch="357"/></w:sectPr>
  </w:body>
</w:document>
'''

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
    <w:pPrDefault><w:pPr></w:pPr></w:pPrDefault>
  </w:docDefaults>
</w:styles>'''

OUT = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx/repro_item6_wrap.docx"
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CONTENT_TYPES)
    z.writestr("_rels/.rels", REL_XML)
    z.writestr("word/_rels/document.xml.rels", DOC_REL_XML)
    z.writestr("word/document.xml", DOC_XML)
    z.writestr("word/styles.xml", STYLES_XML)
print(f"Wrote {OUT}")
