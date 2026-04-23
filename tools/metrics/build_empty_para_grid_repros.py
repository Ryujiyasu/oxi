"""Minimal repros for Word's empty-paragraph grid-snap behavior.

Hypothesis: Word skips grid-snap for empty paragraphs (no runs / bookmark only),
using font metrics instead. This causes d77a +2.5pt drift (Oxi 18pt grid-snap
vs Word 15.5pt no-snap on fs=10.5).

Repros (varying font size & grid pitch):
  G1: Normal sz=21 (10.5pt), docGrid 360tw (18pt). Predict: Word gap 15.5pt.
  G2: Normal sz=21 (10.5pt), docGrid 300tw (15pt). Predict: Word gap 15.5pt
      (if no-snap) or 15pt (if snap).
  G3: Normal sz=24 (12pt),   docGrid 360tw (18pt). Predict: Word gap ~18pt.
  G4: Normal sz=21 (10.5pt), docGrid disabled.     Predict: gap 15.5pt (no grid).
  G5: Normal sz=21 (10.5pt), docGrid 360tw,
      para 1 HAS text (not empty).                Predict: gap 18pt (snapped).
  G6: Normal sz=21 (10.5pt), docGrid 360tw,
      para 1 empty with explicit pStyle.          Predict: ? (test pStyle effect)
"""
import os, zipfile, shutil
from pathlib import Path

OUT = Path(__file__).parent / "empty_para_grid_repro"
OUT.mkdir(exist_ok=True)


def make_docx(name: str, *, normal_sz_halfpt: int = 21, line_pitch_tw: int = 360,
              grid_enabled: bool = True, para1_has_text: bool = False,
              para1_with_pstyle: bool = False,
              para1_ppr_rpr_sz: int = 0,
              para1_ppr_rpr_gothic: bool = False,
              para1_run_rpr_sz: int = 0,
              use_fe_layout: bool = False,
              normal_szcs_halfpt: int = 0,
              normal_kern: int = 0,
              include_para1_szcs: bool = True,   # include szCs in pPr.rPr
              include_para1_jc: bool = True,     # include jc in pPr
              include_para1_bookmark: bool = True,
              ) -> Path:
    """Build a tiny docx testing one variable."""
    # Base XMLs
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

    fe_tag = '<w:useFELayout/>' if use_fe_layout else ''
    settings = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat>{fe_tag}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''

    szcs_val = normal_szcs_halfpt or normal_sz_halfpt
    kern_tag = f'<w:kern w:val="{normal_kern}"/>' if normal_kern else ''
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
    <w:rPr>{kern_tag}<w:sz w:val="{normal_sz_halfpt}"/><w:szCs w:val="{szcs_val}"/></w:rPr>
  </w:style>
</w:styles>'''

    # docGrid / sectPr
    if grid_enabled:
        doc_grid = f'<w:docGrid w:type="lines" w:linePitch="{line_pitch_tw}"/>'
    else:
        doc_grid = ''

    sect_pr = f'''<w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      {doc_grid}
    </w:sectPr>'''

    # Optional pPr.rPr block
    ppr_rpr_parts = []
    if para1_ppr_rpr_gothic:
        ppr_rpr_parts.append('<w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/>')
    if para1_ppr_rpr_sz:
        ppr_rpr_parts.append(f'<w:sz w:val="{para1_ppr_rpr_sz}"/>')
        if include_para1_szcs:
            ppr_rpr_parts.append(f'<w:szCs w:val="{para1_ppr_rpr_sz}"/>')
    ppr_rpr = f'<w:rPr>{"".join(ppr_rpr_parts)}</w:rPr>' if ppr_rpr_parts else ''
    jc_tag = '<w:jc w:val="right"/>' if include_para1_jc else ''
    bookmark_xml = '<w:bookmarkStart w:id="0" w:name="b0"/><w:bookmarkEnd w:id="0"/>' if include_para1_bookmark else ''

    # Run rPr for text case
    run_rpr = ''
    if para1_run_rpr_sz:
        run_rpr = f'<w:rPr><w:sz w:val="{para1_run_rpr_sz}"/><w:szCs w:val="{para1_run_rpr_sz}"/></w:rPr>'

    # Para 1: empty (bookmark-only) or has text
    if para1_has_text:
        para1 = f'<w:p><w:pPr>{ppr_rpr}{jc_tag}</w:pPr><w:r>{run_rpr}<w:t>x</w:t></w:r></w:p>'
    elif para1_with_pstyle:
        para1 = f'<w:p><w:pPr><w:pStyle w:val="a"/>{ppr_rpr}{jc_tag}</w:pPr>{bookmark_xml}</w:p>'
    else:
        para1 = f'<w:p><w:pPr>{ppr_rpr}{jc_tag}</w:pPr>{bookmark_xml}</w:p>'

    # Para 2: a title with sz=28 (14pt), to measure gap
    para2 = '<w:p><w:pPr><w:spacing w:line="420" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>PARA_TWO</w:t></w:r></w:p>'

    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {para1}
    {para2}
    {sect_pr}
  </w:body>
</w:document>'''

    out_path = OUT / f"{name}.docx"
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels_root)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/settings.xml", settings)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", document)
    return out_path


if __name__ == "__main__":
    cases = [
        # Original G1-G6 (reference, keep as-is)
        ("G1_sz21_grid360",    dict(normal_sz_halfpt=21, line_pitch_tw=360)),
        ("G2_sz21_grid300",    dict(normal_sz_halfpt=21, line_pitch_tw=300)),
        ("G3_sz24_grid360",    dict(normal_sz_halfpt=24, line_pitch_tw=360)),
        ("G4_sz21_no_grid",    dict(normal_sz_halfpt=21, grid_enabled=False)),
        ("G5_sz21_grid360_textp1", dict(normal_sz_halfpt=21, line_pitch_tw=360, para1_has_text=True)),
        ("G6_sz21_grid360_pstyle", dict(normal_sz_halfpt=21, line_pitch_tw=360, para1_with_pstyle=True)),
        # Axis 1: grid sweep at fs=10.5 (sz=21)
        ("H01_sz21_grid240",   dict(normal_sz_halfpt=21, line_pitch_tw=240)),   # 12pt
        ("H02_sz21_grid280",   dict(normal_sz_halfpt=21, line_pitch_tw=280)),   # 14pt
        ("H03_sz21_grid320",   dict(normal_sz_halfpt=21, line_pitch_tw=320)),   # 16pt
        ("H04_sz21_grid400",   dict(normal_sz_halfpt=21, line_pitch_tw=400)),   # 20pt
        ("H05_sz21_grid480",   dict(normal_sz_halfpt=21, line_pitch_tw=480)),   # 24pt
        ("H06_sz21_grid600",   dict(normal_sz_halfpt=21, line_pitch_tw=600)),   # 30pt
        ("H07_sz21_grid720",   dict(normal_sz_halfpt=21, line_pitch_tw=720)),   # 36pt
        # Axis 2: fs sweep at grid=18pt (linePitch=360)
        ("H10_sz18_grid360",   dict(normal_sz_halfpt=18, line_pitch_tw=360)),   # fs=9
        ("H11_sz20_grid360",   dict(normal_sz_halfpt=20, line_pitch_tw=360)),   # fs=10
        ("H12_sz22_grid360",   dict(normal_sz_halfpt=22, line_pitch_tw=360)),   # fs=11
        ("H13_sz28_grid360",   dict(normal_sz_halfpt=28, line_pitch_tw=360)),   # fs=14
        ("H14_sz32_grid360",   dict(normal_sz_halfpt=32, line_pitch_tw=360)),   # fs=16
        ("H15_sz36_grid360",   dict(normal_sz_halfpt=36, line_pitch_tw=360)),   # fs=18
        # Axis 3: natural (no grid) at various fs — baseline metrics
        ("H20_sz24_no_grid",   dict(normal_sz_halfpt=24, grid_enabled=False)),  # fs=12
        ("H21_sz28_no_grid",   dict(normal_sz_halfpt=28, grid_enabled=False)),  # fs=14
        ("H22_sz32_no_grid",   dict(normal_sz_halfpt=32, grid_enabled=False)),  # fs=16
        ("H23_sz36_no_grid",   dict(normal_sz_halfpt=36, grid_enabled=False)),  # fs=18
        ("H24_sz20_no_grid",   dict(normal_sz_halfpt=20, grid_enabled=False)),  # fs=10
        # Axis 4: diagonal probes — fs close to grid_pitch
        ("H30_sz24_grid300",   dict(normal_sz_halfpt=24, line_pitch_tw=300)),   # fs=12, grid=15
        ("H31_sz28_grid480",   dict(normal_sz_halfpt=28, line_pitch_tw=480)),   # fs=14, grid=24
        ("H32_sz32_grid480",   dict(normal_sz_halfpt=32, line_pitch_tw=480)),   # fs=16, grid=24
        # Axis 5: pPr.rPr override (mirror d77a p3 setup)
        # Normal sz=21 (fs=10.5), para 1 has pPr.rPr.sz override (fs=12 via sz=24)
        # Predict formula: fs=12 grid=18 → gap=16. But d77a p3 observed=18.
        ("H40_norm21_pprrpr24", dict(normal_sz_halfpt=21, line_pitch_tw=360, para1_ppr_rpr_sz=24)),
        ("H41_norm21_pprrpr24_gothic", dict(normal_sz_halfpt=21, line_pitch_tw=360, para1_ppr_rpr_sz=24, para1_ppr_rpr_gothic=True)),
        # Sanity: same config as H40 but with text, to see if empty vs non-empty matters
        ("H42_norm21_pprrpr24_text", dict(normal_sz_halfpt=21, line_pitch_tw=360, para1_ppr_rpr_sz=24, para1_has_text=True)),
        # pPr.rPr at a smaller size (fs=12→fs=10.5 case already covered by G1)
        # Try large override fs=14, fs=16 via pPr.rPr
        ("H43_norm21_pprrpr28", dict(normal_sz_halfpt=21, line_pitch_tw=360, para1_ppr_rpr_sz=28)),   # fs=14
        ("H44_norm21_pprrpr32", dict(normal_sz_halfpt=21, line_pitch_tw=360, para1_ppr_rpr_sz=32)),   # fs=16
        # Axis 6: run rPr override with text (empirically relevant?)
        ("H50_norm21_text_runsz24", dict(normal_sz_halfpt=21, line_pitch_tw=360, para1_has_text=True, para1_run_rpr_sz=24)),
        ("H51_norm21_text_runsz28", dict(normal_sz_halfpt=21, line_pitch_tw=360, para1_has_text=True, para1_run_rpr_sz=28)),
        # Axis 7: useFELayout (Far East Layout compat) — mimic d77a setting
        # H60: replicate d77a p3 exactly — Normal sz=21 szCs=24 kern=2, pPr.rPr sz=24 MS Gothic, useFELayout
        ("H60_d77a_p3_exact", dict(
            normal_sz_halfpt=21, normal_szcs_halfpt=24, normal_kern=2,
            line_pitch_tw=360,
            para1_ppr_rpr_sz=24, para1_ppr_rpr_gothic=True,
            use_fe_layout=True)),
        # H61: isolate useFELayout effect (otherwise == H41)
        ("H61_fe_layout_gothic", dict(
            normal_sz_halfpt=21,
            line_pitch_tw=360,
            para1_ppr_rpr_sz=24, para1_ppr_rpr_gothic=True,
            use_fe_layout=True)),
        # H62: isolate kern effect (otherwise == H41)
        ("H62_kern_gothic", dict(
            normal_sz_halfpt=21, normal_kern=2,
            line_pitch_tw=360,
            para1_ppr_rpr_sz=24, para1_ppr_rpr_gothic=True)),
        # H63: isolate szCs mismatch (otherwise == H41)
        ("H63_szcs24_gothic", dict(
            normal_sz_halfpt=21, normal_szcs_halfpt=24,
            line_pitch_tw=360,
            para1_ppr_rpr_sz=24, para1_ppr_rpr_gothic=True)),
        # H64: useFELayout with MS Mincho (no gothic)
        ("H64_fe_mincho", dict(
            normal_sz_halfpt=21,
            line_pitch_tw=360,
            para1_ppr_rpr_sz=24,
            use_fe_layout=True)),
        # H65: useFELayout, empty, Normal sz=24 (like G3 but with FELayout)
        ("H65_fe_sz24", dict(
            normal_sz_halfpt=24,
            line_pitch_tw=360,
            use_fe_layout=True)),
        # H66: baseline (G1-equiv) with useFELayout, to confirm fs=10.5 still gives 15.5
        ("H66_fe_sz21", dict(
            normal_sz_halfpt=21,
            line_pitch_tw=360,
            use_fe_layout=True)),
        # H67: strip szCs/jc/bookmark from H60 to match d77a p3 exactly
        ("H67_d77a_p3_strip", dict(
            normal_sz_halfpt=21, normal_szcs_halfpt=24, normal_kern=2,
            line_pitch_tw=360,
            para1_ppr_rpr_sz=24, para1_ppr_rpr_gothic=True,
            use_fe_layout=True,
            include_para1_szcs=False,
            include_para1_jc=False,
            include_para1_bookmark=False)),
        # H68: strip only szCs from H60
        ("H68_no_szcs", dict(
            normal_sz_halfpt=21, normal_szcs_halfpt=24, normal_kern=2,
            line_pitch_tw=360,
            para1_ppr_rpr_sz=24, para1_ppr_rpr_gothic=True,
            use_fe_layout=True,
            include_para1_szcs=False)),
        # H69: strip only bookmark from H60
        ("H69_no_bookmark", dict(
            normal_sz_halfpt=21, normal_szcs_halfpt=24, normal_kern=2,
            line_pitch_tw=360,
            para1_ppr_rpr_sz=24, para1_ppr_rpr_gothic=True,
            use_fe_layout=True,
            include_para1_bookmark=False)),
    ]
    for name, kwargs in cases:
        p = make_docx(name, **kwargs)
        print(f"created: {p}")
