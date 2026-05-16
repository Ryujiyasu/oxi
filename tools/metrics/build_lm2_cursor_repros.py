"""Minimal repros for linesAndChars (LM2) mode cursor advance behavior.

Hypothesis (R56 candidate): in `<w:docGrid w:type="linesAndChars"
w:linePitch=N/>` mode, paragraphs ALWAYS start at a grid line boundary
in Word (cursor snaps to next cell at paragraph entry). Oxi's
is_lm2_single path at mod.rs:4414 uses `k = floor((cur-margin)/pitch)`
without snapping, leading to mid-cell paragraph entries and ~1pt
advance instead of one full cell pitch.

d1e8 (linePitch=292tw=14.6pt) shows: pi=30 enters mid-cell at y=128,
advance only 1pt to y=129. Word renders wi=31->wi=32 with 14.5pt advance.

Repros:
  M1: linePitch=292, all Single 10.5pt. 4 simple paragraphs.
      Expect: each advance = 14.5pt (or 14.6 = pitch).
  M2: linePitch=292, mix Single + 1.5x (line=360 auto) empty.
      d1e8 pattern. Empty 1.5x → 2 cells? 1 cell?
  M3: linePitch=292, paragraph with firstLine indent 84pt (firstLine=1680).
      d1e8 wi=31 paragraph type.
  M4: linePitch=357 (matches 1ec1). Expect aligned advances.
  M5: linePitch=292, multi-run TEXT + multiple short runs.
      Tests d1e8 multi-run pattern.
  M6: linePitch=292, 1.5x empty followed by indented TEXT. d1e8 wi=30->wi=31->wi=32.
"""
import os, zipfile
from pathlib import Path

OUT = Path(__file__).parent / "lm2_cursor_repro"
OUT.mkdir(exist_ok=True)


def make_docx(name: str, *, body_xml: str,
              normal_sz_halfpt: int = 21,
              line_pitch_tw: int = 292,
              grid_type: str = "linesAndChars",
              char_space_tw: int | None = None,
              ) -> Path:
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''
    rels_root = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''
    settings = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''
    styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/></w:rPr></w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="a">
    <w:name w:val="Normal"/>
    <w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
    <w:rPr><w:sz w:val="{normal_sz_halfpt}"/><w:szCs w:val="{normal_sz_halfpt}"/></w:rPr>
  </w:style>
</w:styles>'''
    char_space_attr = f' w:charSpace="{char_space_tw}"' if char_space_tw else ''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body_xml}
<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1701" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>
  <w:cols w:space="425"/>
  <w:docGrid w:type="{grid_type}" w:linePitch="{line_pitch_tw}"{char_space_attr}/>
</w:sectPr>
</w:body>
</w:document>'''
    out_path = OUT / f"{name}.docx"
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', rels_root)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/settings.xml', settings)
    return out_path


P_TEXT = '<w:p><w:r><w:t>あいうえお</w:t></w:r></w:p>'
P_TEXT_INDENT = '<w:p><w:pPr><w:ind w:firstLineChars="800" w:firstLine="1680"/></w:pPr><w:r><w:t>あいうえお</w:t></w:r></w:p>'
P_EMPTY = '<w:p/>'
P_EMPTY_15X = '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:p>'
P_MULTIRUN = '<w:p><w:r><w:t>あ</w:t></w:r><w:r><w:t> </w:t></w:r><w:r><w:t>い</w:t></w:r><w:r><w:t> </w:t></w:r><w:r><w:t>う</w:t></w:r></w:p>'

variants = {
    "M1_pitch292_4text":        P_TEXT * 4,
    "M2_pitch292_text_then_15x_empty": P_TEXT + P_EMPTY_15X + P_TEXT + P_EMPTY_15X + P_TEXT,
    "M3_pitch292_indent_text":  P_TEXT_INDENT * 4,
    "M4_pitch357_4text":        P_TEXT * 4,  # different pitch
    "M5_pitch292_multirun":     P_MULTIRUN * 4,
    "M6_d1e8_pattern":          P_EMPTY_15X + P_TEXT_INDENT + P_TEXT_INDENT + P_TEXT_INDENT,
}

results_paths = {}
for name, body in variants.items():
    pitch = 357 if name.startswith("M4") else 292
    p = make_docx(name, body_xml=body, line_pitch_tw=pitch)
    results_paths[name] = p
    print(f"  {p}")
print(f"\nbuilt {len(variants)} repros in {OUT}")
