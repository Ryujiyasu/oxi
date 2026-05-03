"""Day 6: indent matrix to pin Word's table_x / first-line position rule.

Variants from baseline (hang-indent like 1636 item 6):
- v0_baseline: left=5775 leftChars=150 hanging=5460 hangingChars=2600 (full 1636 setup)
- v1_no_leftChars: left=5775 hanging=5460 (twip only, no chars)
- v2_no_left:    leftChars=150 hangingChars=2600 (chars only)
- v3_small_left: left=315 leftChars=15 hanging=0 (small indent)
- v4_no_hanging: left=5775 leftChars=150 (no hang)
- v5_only_hangingChars: hangingChars=2600 (chars hanging only, no twip)
- v6_only_hanging: hanging=5460 (twip hanging only, no chars)
- v7_huge_left: left=10000 leftChars=200 hanging=10000 hangingChars=2000

For each, render text "ABCDEあいうえお" with simple cell, measure Word + Oxi first char x.
"""
import zipfile, os

CTYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
REL = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_REL = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>'''

# Use 1636-like structure: tblW=9628, gridCol=9628, no tblCellMar, no tblInd (to isolate indent effect)
def build_doc(ind_attrs, text="ABCDEあいうえお"):
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

VARIANTS = {
    "v0_full":              'w:leftChars="150" w:left="5775" w:hangingChars="2600" w:hanging="5460"',
    "v1_no_leftChars":      'w:left="5775" w:hanging="5460"',
    "v2_no_left":           'w:leftChars="150" w:hangingChars="2600"',
    "v3_small_left":        'w:leftChars="15" w:left="315"',
    "v4_no_hanging":        'w:leftChars="150" w:left="5775"',
    "v5_chars_hanging_only": 'w:hangingChars="2600"',
    "v6_twip_hanging_only": 'w:hanging="5460"',
    "v7_huge_left":         'w:leftChars="200" w:left="10000" w:hangingChars="2000" w:hanging="10000"',
    "v8_zero_indent":       '',  # no ind at all
}

OUT_DIR = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
for name, ind in VARIANTS.items():
    path = f"{OUT_DIR}/indent_matrix_{name}.docx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CTYPES)
        z.writestr("_rels/.rels", REL)
        z.writestr("word/_rels/document.xml.rels", DOC_REL)
        z.writestr("word/document.xml", build_doc(ind))
        z.writestr("word/styles.xml", STYLES)
    print(f"Wrote {path}")
print(f"\nTotal: {len(VARIANTS)} variants")
