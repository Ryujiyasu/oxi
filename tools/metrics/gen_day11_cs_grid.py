"""Day 11: isolate character_spacing (cs) effect from docGrid effect.

Variants:
- v0_baseline: no cs, no docGrid (pure font width)
- v1_cs_neg9: cs=-9 only, no docGrid
- v2_cs_neg20: cs=-20 only, no docGrid
- v3_grid_only: no cs, with docGrid linesAndChars linePitch=272
- v4_grid_cs_neg9: cs=-9 + docGrid (1636-like)
- v5_grid_cs_neg1: cs=-1 + docGrid

For each, measure Word per-char x positions to derive per-char advance.
"""
import zipfile

CTYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
REL = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_REL = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>'''

# 20 fullwidth CJK chars at 10.5pt (sz=21)
TEXT = "あいうえおかきくけこさしすせそたちつてと"

def build_doc(cs_val=None, with_grid=False):
    cs_xml = f'<w:rPr><w:spacing w:val="{cs_val}"/></w:rPr>' if cs_val is not None else ''
    grid_xml = '<w:docGrid w:type="linesAndChars" w:linePitch="272"/>' if with_grid else ''
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r>{cs_xml}<w:t>{TEXT}</w:t></w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="851" w:right="851" w:bottom="851" w:left="851" w:gutter="0"/>{grid_xml}</w:sectPr>
  </w:body>
</w:document>
'''

VARIANTS = [
    ("d11_v0_baseline", None, False),
    ("d11_v1_cs_neg9", -9, False),
    ("d11_v2_cs_neg20", -20, False),
    ("d11_v3_grid_only", None, True),
    ("d11_v4_grid_cs_neg9", -9, True),
    ("d11_v5_grid_cs_neg1", -1, True),
]
OUT_DIR = "c:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"
for name, cs, grid in VARIANTS:
    path = f"{OUT_DIR}/{name}.docx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CTYPES)
        z.writestr("_rels/.rels", REL)
        z.writestr("word/_rels/document.xml.rels", DOC_REL)
        z.writestr("word/document.xml", build_doc(cs, grid))
        z.writestr("word/styles.xml", STYLES)
    print(f"Wrote {path}")
print(f"\nTotal: {len(VARIANTS)} variants")
