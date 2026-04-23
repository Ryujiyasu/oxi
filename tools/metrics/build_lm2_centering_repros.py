"""Minimal repros to determine Word's LM2 single-cell glyph positioning rule.

Question: For font_size ≤ grid_pitch, does Word:
  A) Center glyph top at (pitch - font_size)/2 above line_top
  B) Place glyph top at line_top (no centering)
  C) Something else (ascent-based, etc.)

Hypothesis from d77a: Word uses rule B (glyph top at line top ~+1pt).
Memory says (2026-04-04/16) Word centers for LM2 multi-cell. Single-cell
unconfirmed.

Repros C1-C6: MS Mincho, single line, one para only.
  C1: fs=10.5pt (21 half), grid 360tw (18pt) — d77a case
  C2: fs=12.0pt (24 half), grid 360tw (18pt)
  C3: fs=14.0pt (28 half), grid 360tw (18pt)
  C4: fs=10.5pt, grid 300tw (15pt) — pitch=fs+4.5
  C5: fs=10.5pt, grid disabled (LM0) — baseline
  C6: fs=10.5pt, grid 360tw, compat=14 (vs default 15)

Expected COM output per repro: first paragraph Range.Information(6) = glyph
top Y. Subtract page_top margin (70.9pt default) to get offset.
"""
import os, zipfile
from pathlib import Path

OUT = Path(__file__).parent / "lm2_centering_repro"
OUT.mkdir(exist_ok=True)


def make_docx(name: str, *, fs_halfpt: int = 21, line_pitch_tw: int = 360,
              grid_enabled: bool = True, compat_mode: int = 15) -> Path:
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

    rels_doc = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

    styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="MS Mincho" w:hAnsi="MS Mincho" w:eastAsia="MS Mincho" w:cs="Times New Roman"/>
        <w:sz w:val="{fs_halfpt}"/>
        <w:szCs w:val="{fs_halfpt}"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr/>
    <w:rPr>
      <w:rFonts w:ascii="MS Mincho" w:hAnsi="MS Mincho" w:eastAsia="MS Mincho"/>
      <w:sz w:val="{fs_halfpt}"/>
      <w:szCs w:val="{fs_halfpt}"/>
    </w:rPr>
  </w:style>
</w:styles>'''

    settings = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat_mode}"/>
  </w:compat>
</w:settings>'''

    grid_xml = f'<w:docGrid w:type="lines" w:linePitch="{line_pitch_tw}"/>' if grid_enabled else '<w:docGrid w:type="default"/>'

    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Test line AAAAAA</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Second line BBBBBB</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397" w:gutter="0"/>
      {grid_xml}
    </w:sectPr>
  </w:body>
</w:document>'''

    out_path = OUT / f"{name}.docx"
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels_root)
        z.writestr("word/_rels/document.xml.rels", rels_doc)
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
    return out_path


def make_docx_font(name: str, *, font: str, fs_halfpt: int, line_pitch_tw: int = 360) -> Path:
    """Variant with custom font."""
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
    rels_doc = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''
    styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="Times New Roman"/>
        <w:sz w:val="{fs_halfpt}"/>
        <w:szCs w:val="{fs_halfpt}"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr/>
    <w:rPr>
      <w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}"/>
      <w:sz w:val="{fs_halfpt}"/>
      <w:szCs w:val="{fs_halfpt}"/>
    </w:rPr>
  </w:style>
</w:styles>'''
    settings = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Test line AAAAAA</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second line BBBBBB</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397" w:gutter="0"/>
      <w:docGrid w:type="lines" w:linePitch="{line_pitch_tw}"/>
    </w:sectPr>
  </w:body>
</w:document>'''
    out_path = OUT / f"{name}.docx"
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels_root)
        z.writestr("word/_rels/document.xml.rels", rels_doc)
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
    return out_path


if __name__ == "__main__":
    variants = [
        ("C1_fs10_5_pitch18", dict(fs_halfpt=21, line_pitch_tw=360)),
        ("C2_fs12_pitch18",   dict(fs_halfpt=24, line_pitch_tw=360)),
        ("C3_fs14_pitch18",   dict(fs_halfpt=28, line_pitch_tw=360)),
        ("C4_fs10_5_pitch15", dict(fs_halfpt=21, line_pitch_tw=300)),
        ("C5_fs10_5_nogrid",  dict(fs_halfpt=21, grid_enabled=False)),
        ("C6_fs10_5_compat14",dict(fs_halfpt=21, line_pitch_tw=360, compat_mode=14)),
    ]
    for name, kw in variants:
        p = make_docx(name, **kw)
        print(f"Built: {p}")
    # Additional font variants
    font_variants = [
        ("C7_mincho_fs10_5", dict(font="MS Mincho", fs_halfpt=21)),
        ("C8_gothic_fs10_5", dict(font="MS Gothic", fs_halfpt=21)),
        ("C9_meiryo_fs10_5", dict(font="Meiryo",    fs_halfpt=21)),
        ("C10_gothic_fs12",  dict(font="MS Gothic", fs_halfpt=24)),
        ("C11_meiryo_fs12",  dict(font="Meiryo",    fs_halfpt=24)),
    ]
    for name, kw in font_variants:
        p = make_docx_font(name, **kw)
        print(f"Built: {p}")
