"""R92 minimal repro: empty paragraph height in lineRule=exact mode.

Builds a 5-paragraph docx isolating the R90/R91 hypothesis:
- Para 1: text "Line 1" at sz=16 (8pt CJK), lineRule=exact line=240 (12pt)
- Para 2: EMPTY paragraph with same pPr
- Para 3: text "Line 3" at sz=16 same pPr
- Para 4: EMPTY paragraph (lineRule=auto + snap=0 + sz=20 (10pt) — R90 config)
- Para 5: text "Line 5" with R90 config

Expected per Word: 
- Para 1 → 3 step = 24pt (2 × 12pt exact)
- Para 3 → 5 step = ~13.55pt avg (R90 measured pattern)

Oxi-side measurement: cargo run --example layout_json on this fixture
+ COM measurement gives delta to pinpoint over-estimate component.
"""
from __future__ import annotations
import zipfile
import pathlib

OUT = pathlib.Path(__file__).parent / "minimal_empty_para.docx"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Century" w:eastAsia="MS Mincho" w:hAnsi="Century"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr/></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
"""

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>
"""

# Page setup: A4 portrait, 1in margins
DOC = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <!-- Para 1: 8pt CJK with lineRule=exact line=240 (12pt) -->
    <w:p>
      <w:pPr>
        <w:spacing w:line="240" w:lineRule="exact"/>
        <w:rPr><w:sz w:val="16"/><w:szCs w:val="20"/></w:rPr>
      </w:pPr>
      <w:r><w:rPr><w:sz w:val="16"/><w:szCs w:val="20"/></w:rPr><w:t>Line 1 (R91 config: sz=8pt lineRule=exact line=12pt)</w:t></w:r>
    </w:p>
    <!-- Para 2: EMPTY with same pPr -->
    <w:p>
      <w:pPr>
        <w:spacing w:line="240" w:lineRule="exact"/>
        <w:rPr><w:sz w:val="16"/><w:szCs w:val="20"/></w:rPr>
      </w:pPr>
    </w:p>
    <!-- Para 3: text resumes -->
    <w:p>
      <w:pPr>
        <w:spacing w:line="240" w:lineRule="exact"/>
        <w:rPr><w:sz w:val="16"/><w:szCs w:val="20"/></w:rPr>
      </w:pPr>
      <w:r><w:rPr><w:sz w:val="16"/><w:szCs w:val="20"/></w:rPr><w:t>Line 3 (after empty para 2)</w:t></w:r>
    </w:p>
    <!-- Para 4: EMPTY R90-style: snap=0 + lineRule=auto + sz=10pt -->
    <w:p>
      <w:pPr>
        <w:snapToGrid w:val="0"/>
        <w:spacing w:line="240" w:lineRule="auto"/>
        <w:rPr><w:sz w:val="20"/><w:szCs w:val="22"/></w:rPr>
      </w:pPr>
    </w:p>
    <!-- Para 5: text resumes R90-style -->
    <w:p>
      <w:pPr>
        <w:snapToGrid w:val="0"/>
        <w:spacing w:line="240" w:lineRule="auto"/>
        <w:rPr><w:sz w:val="20"/><w:szCs w:val="22"/></w:rPr>
      </w:pPr>
      <w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="22"/></w:rPr><w:t>Line 5 (R90 config: sz=10pt snap=0 line=auto)</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
    </w:sectPr>
  </w:body>
</w:document>
"""

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CONTENT_TYPES)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("word/_rels/document.xml.rels", DOC_RELS)
    z.writestr("word/styles.xml", STYLES)
    z.writestr("word/settings.xml", SETTINGS)
    z.writestr("word/document.xml", DOC)
print(f"wrote {OUT} ({OUT.stat().st_size} bytes)")
