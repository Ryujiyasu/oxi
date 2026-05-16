"""COM-confirm Word's `<w:p><w:pPr><w:spacing w:line="0" w:lineRule="atLeast"/></w:pPr></w:p>`
behavior for EMPTY paragraphs.

R55 fixed line=0 atLeast for non-empty paragraphs (L1-L8). d1e8 has
EMPTY paragraphs with line=0 atLeast between text paragraphs (wi=36-44
on page 2). R55 unmasked a -12.5pt-per-empty cascade — Oxi may be
using natural line height (correct per L1-L8) but Word uses something
different for empties.

Variants:
  E1: 4 text + interleaved empty (4 empties), all line=0 atLeast,
      doc default size 10.5pt (matches d1e8 wi=36-44 pattern)
  E2: same but EMPTY paragraphs have NO line spacing (default Single).
      Tests whether the line=0 atLeast on empty is the issue.
  E3: 1 text + 1 empty (line=0 atLeast) + 1 text. Bare minimum.
  E4: same as E3 but empty has w:rPr with sz=22 (11pt).
  E5: 4 empties at line=0 atLeast then 1 text. Tests pure-empty cluster.

All use docGrid linePitch=360tw (18pt).
"""
import os, zipfile
from pathlib import Path

OUT = Path(__file__).parent / "line0_atleast_empty_repro"
OUT.mkdir(exist_ok=True)


def make_docx(name: str, *, body_xml: str,
              normal_sz_halfpt: int = 21,
              line_pitch_tw: int = 360,
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
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body_xml}
<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1985" w:right="1701" w:bottom="1985" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>
  <w:cols w:space="425"/>
  <w:docGrid w:type="lines" w:linePitch="{line_pitch_tw}"/>
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


SP_ATLEAST_0 = '<w:spacing w:line="0" w:lineRule="atLeast"/>'

P_TEXT_AL0 = f'<w:p><w:pPr>{SP_ATLEAST_0}</w:pPr><w:r><w:t>テキスト</w:t></w:r></w:p>'
P_TEXT_NO  = '<w:p><w:r><w:t>テキスト</w:t></w:r></w:p>'
P_EMPTY_AL0 = f'<w:p><w:pPr>{SP_ATLEAST_0}</w:pPr></w:p>'
P_EMPTY_NO  = '<w:p/>'
P_EMPTY_AL0_SZ22 = f'<w:p><w:pPr>{SP_ATLEAST_0}<w:rPr><w:sz w:val="22"/></w:rPr></w:pPr></w:p>'

variants = {
    "E1_text_empty_alternate_all_AL0": (P_TEXT_AL0 + P_EMPTY_AL0) * 4 + P_TEXT_AL0,
    "E2_text_AL0_empty_default":       (P_TEXT_AL0 + P_EMPTY_NO ) * 4 + P_TEXT_AL0,
    "E3_text_empty_text_AL0":          P_TEXT_AL0 + P_EMPTY_AL0 + P_TEXT_AL0,
    "E4_text_empty_sz22_text":         P_TEXT_AL0 + P_EMPTY_AL0_SZ22 + P_TEXT_AL0,
    "E5_4empty_then_text":             P_EMPTY_AL0 * 4 + P_TEXT_AL0,
    "E6_text_no_spacing_empty_AL0":    (P_TEXT_NO  + P_EMPTY_AL0 ) * 4 + P_TEXT_NO,
}

for name, body in variants.items():
    p = make_docx(name, body_xml=body)
    print(f"  {p}")
print(f"\nbuilt {len(variants)} repros in {OUT}")
