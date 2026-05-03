"""Day 7: line-2 wrap + bullet variants to pin Word's first-line-indent rule."""
import zipfile, os

CTYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''
REL = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_REL = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>'''
NUMBERING = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="□"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>'''

# Long text that WILL wrap (50+ CJK chars at 12pt = 600pt natural in 480pt cell)
LONG_TEXT = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをんがぎぐげご"

def build_doc(ind_attrs, text=LONG_TEXT, numPr=False):
    np = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if numPr else ''
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:type="dxa" w:w="9628"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="9628"/></w:tblGrid>
      <w:tr><w:trHeight w:val="2000"/>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9628"/></w:tcPr>
          <w:p>
            <w:pPr>
              <w:spacing w:line="280" w:lineRule="exact"/>
              {np}
              <w:ind {ind_attrs}/>
            </w:pPr>
            <w:r><w:t xml:space="preserve">{text}</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="851" w:right="1134" w:bottom="142" w:left="1134" w:gutter="0"/><w:docGrid w:type="linesAndChars" w:linePitch="357"/></w:sectPr>
  </w:body>
</w:document>
'''

# Variants — focus on hanging-indent + line wrap to see line 2 position
VARIANTS = [
    ("d7_v0_hang_only", 'w:left="600" w:hanging="600"', False),  # left=30pt, hanging=30pt → first_line=0, line 2 at 30pt
    ("d7_v1_hang_offset", 'w:left="600" w:hanging="300"', False),  # first_line=15pt, line 2 at 30pt
    ("d7_v2_no_hang_left30", 'w:left="600"', False),  # left=30, no hang → both lines at 30pt
    ("d7_v3_huge_hang_like_1636", 'w:leftChars="150" w:left="5775" w:hangingChars="2600" w:hanging="5460"', False),  # 1636 item 6
    ("d7_v4_bullet", 'w:left="420" w:hanging="420"', True),  # bullet list with hanging
    ("d7_v5_no_indent_bullet", '', True),  # bullet list, no indent
]

OUT_DIR = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
for name, ind, num in VARIANTS:
    path = f"{OUT_DIR}/{name}.docx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CTYPES)
        z.writestr("_rels/.rels", REL)
        z.writestr("word/_rels/document.xml.rels", DOC_REL)
        z.writestr("word/document.xml", build_doc(ind, numPr=num))
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)
    print(f"Wrote {path}")
print(f"\nTotal: {len(VARIANTS)} variants")
