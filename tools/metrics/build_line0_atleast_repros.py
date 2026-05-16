"""Minimal repros for Word's `line=0 atLeast` line-height behavior.

Hypothesis (R55 candidate, 2026-05-17): Oxi's `atLeast` branch at
crates/oxidocs-core/src/layout/mod.rs:5479-5489 applies grid-snap to
the natural line height, then takes max(snapped, val). For val=0
(`<w:spacing w:line="0" w:lineRule="atLeast"/>`), this becomes
grid_snap(natural). e201 (mean_iou=0.3346) and d1e8 (0.6376) — the
ONLY 2 baseline docs using line=0/atLeast — both show large
accumulating Y drift (e201 +1.9 -> +87.7pt monotonic over 11 paras).

Hypothesis: Word treats `line=0 atLeast` as "use natural line height
without grid-snap" — the snap is a misapplication. This script builds
variants to COM-confirm:

  L1: 10.5pt MS Mincho, line=0 atLeast (4 paragraphs)
  L2: 14pt   MS Mincho, line=0 atLeast (4 paragraphs)
  L3: 10pt   MS Mincho, line=0 atLeast (4 paragraphs)
  L4: 12pt   MS Mincho, line=0 atLeast (4 paragraphs)
  L5: 10.5pt MS Mincho, line=240 atLeast (4 paragraphs)  # baseline for non-zero
  L6: 10.5pt MS Mincho, NO spacing element (Single)       # baseline for snap
  L7: 14pt   MS Mincho, NO spacing element (Single)       # baseline for snap
  L8: mixed  (10.5/14/10/12pt) line=0 atLeast              # e201 pattern

docGrid linePitch=360tw (18pt) on all (matches e201/d1e8 sectPr).

Each variant: 4 paragraphs of single-line text, all same size, so the
per-paragraph advance == line_height. COM measures paragraph y; gap =
y[k+1] - y[k] tells us the actual line height Word uses.
"""
import os, zipfile, shutil
from pathlib import Path

OUT = Path(__file__).parent / "line0_atleast_repro"
OUT.mkdir(exist_ok=True)


def make_docx(name: str, *, paragraphs: list[tuple[int, str, str | None]],
              normal_sz_halfpt: int = 21,
              line_pitch_tw: int = 360,
              ) -> Path:
    """Build a tiny docx with arbitrary paragraphs.

    paragraphs: list of (sz_halfpt, text, spacing_xml). spacing_xml is
        the inner `<w:spacing ... />` element or None for no spacing.
    """
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
    <w:rPrDefault>
      <w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/></w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="a">
    <w:name w:val="Normal"/>
    <w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
    <w:rPr><w:sz w:val="{normal_sz_halfpt}"/><w:szCs w:val="{normal_sz_halfpt}"/></w:rPr>
  </w:style>
</w:styles>'''

    para_xmls = []
    for i, (sz, text, spacing_xml) in enumerate(paragraphs):
        ppr_inner = ''
        if spacing_xml:
            ppr_inner += spacing_xml
        if sz:
            ppr_inner += f'<w:rPr><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        ppr = f'<w:pPr>{ppr_inner}</w:pPr>' if ppr_inner else ''
        run_rpr = f'<w:rPr><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>' if sz else ''
        run = f'<w:r>{run_rpr}<w:t xml:space="preserve">{text}</w:t></w:r>' if text else ''
        para_xmls.append(f'<w:p>{ppr}{run}</w:p>')

    body_paras = ''.join(para_xmls)
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body_paras}
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


# Build variants
SP_ATLEAST_0 = '<w:spacing w:line="0" w:lineRule="atLeast"/>'
SP_ATLEAST_240 = '<w:spacing w:line="240" w:lineRule="atLeast"/>'

variants = {
    "L1_105_atleast0": [(21, "あいう", SP_ATLEAST_0)] * 4,
    "L2_14_atleast0":  [(28, "あいう", SP_ATLEAST_0)] * 4,
    "L3_10_atleast0":  [(20, "あいう", SP_ATLEAST_0)] * 4,
    "L4_12_atleast0":  [(24, "あいう", SP_ATLEAST_0)] * 4,
    "L5_105_atleast240": [(21, "あいう", SP_ATLEAST_240)] * 4,
    "L6_105_single":   [(21, "あいう", None)] * 4,
    "L7_14_single":    [(28, "あいう", None)] * 4,
    "L8_mixed_atleast0": [
        (21, "あいう", SP_ATLEAST_0),
        (28, "あいう", SP_ATLEAST_0),
        (20, "あいう", SP_ATLEAST_0),
        (24, "あいう", SP_ATLEAST_0),
    ],
}

for name, paras in variants.items():
    p = make_docx(name, paragraphs=paras)
    print(f"  {p}")
print(f"\nbuilt {len(variants)} repros in {OUT}")
