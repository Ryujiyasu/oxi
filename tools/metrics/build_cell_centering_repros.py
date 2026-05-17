"""Minimal repros for cell text centering hypotheses (Sessions 69-70).

Two compounding mechanisms:

A. Body exact rule bottom-align bug (Session 70)
   Hypothesis: Oxi's text_y_offset_for_line at mod.rs:5832 returns
   `(lh - fs)` for exact rule (bottom-align). Correct for 1ec1 SHAPE
   but Word top-aligns body exact paragraphs (Information(6) y_pg =
   topMargin for first para). 04b88 dy=+6.5pt cluster.

B. Single rule centering uses font_size instead of GDI cell height
   (Session 69, [[session69-cell-text-y-off-bug]])
   Hypothesis: mod.rs:5920 and mod.rs:6791 use fs directly; should use
   fs × 83/64 (or natural lh) for CJK 83/64 fonts. 04b88 dy=+2pt cluster
   in table cells.

Repros (output in `cell_centering_repro/`):

A series — body paragraph, no table:
  A1: exact line=340 (17pt), MS Mincho 10.5pt — replicates 04b88 first para
  A2: exact line=480 (24pt), MS Mincho 10.5pt — bigger lh, larger expected offset
  A3: exact line=440 (22pt), MS Mincho 14pt — matches 1ec1 SHAPE pattern but in body context
  A4: A1 with pPr <w:textAlignment w:val="top"/> — does Word respond?
  A5: A1 with pPr <w:textAlignment w:val="bottom"/> — does Word actually bottom-align?
  A6: Single rule (no exact) Mincho 10.5pt — baseline (mechanism B alone in body)
  A7: exact line=340 Times New Roman 10.5pt — non-CJK font path

B series — single-row single-cell table:
  B1: 1-cell table, Single Mincho 10.5pt, docGrid lines linePitch=360
  B2: 1-cell table, Single Mincho 14pt — bigger fs
  B3: 1-cell table, Single Yu Mincho 10.5pt — different CJK font
  B4: 1-cell table, Single Times New Roman 10.5pt — non-CJK font
  B5: 1-cell table, exact line=480 Mincho 10.5pt — cell + exact rule
  B6: 1-cell table, atLeast line=240 Mincho 10.5pt — cell + atLeast
"""
import os
import zipfile
from pathlib import Path

OUT = Path(__file__).parent / "cell_centering_repro"
OUT.mkdir(exist_ok=True)

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def make_docx(
    name: str,
    *,
    body_xml: str,
    line_pitch_tw: int = 360,
    grid_type: str = "lines",
    east_asia_font: str = "ＭＳ 明朝",
    ascii_font: str = "Century",
    normal_sz_halfpt: int = 21,
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
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="{ascii_font}" w:eastAsia="{east_asia_font}" w:hAnsi="{ascii_font}"/></w:rPr></w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="a">
    <w:name w:val="Normal"/>
    <w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
    <w:rPr><w:sz w:val="{normal_sz_halfpt}"/><w:szCs w:val="{normal_sz_halfpt}"/></w:rPr>
  </w:style>
</w:styles>'''
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {W_NS}>
<w:body>
{body_xml}
<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304" w:header="851" w:footer="992" w:gutter="0"/>
  <w:cols w:space="425"/>
  <w:docGrid w:type="{grid_type}" w:linePitch="{line_pitch_tw}"/>
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


def body_exact(line_tw: int, fs_halfpt: int, text: str = "あいうえお", text_align: str = None) -> str:
    """Body paragraph with explicit exact spacing."""
    ta = f'<w:textAlignment w:val="{text_align}"/>' if text_align else ''
    return (
        '<w:p>'
        '<w:pPr>'
        f'<w:spacing w:line="{line_tw}" w:lineRule="exact"/>'
        f'{ta}'
        '</w:pPr>'
        '<w:r>'
        f'<w:rPr><w:sz w:val="{fs_halfpt}"/><w:szCs w:val="{fs_halfpt}"/></w:rPr>'
        f'<w:t>{text}</w:t>'
        '</w:r>'
        '</w:p>'
    )


def body_single(fs_halfpt: int, text: str = "あいうえお", font_ea: str = None, font_ascii: str = None) -> str:
    """Body paragraph with Single line spacing (no explicit spacing)."""
    rfonts = ''
    if font_ea or font_ascii:
        ea = f' w:eastAsia="{font_ea}"' if font_ea else ''
        ac = f' w:ascii="{font_ascii}" w:hAnsi="{font_ascii}"' if font_ascii else ''
        rfonts = f'<w:rFonts{ea}{ac}/>'
    return (
        '<w:p>'
        '<w:r>'
        f'<w:rPr>{rfonts}<w:sz w:val="{fs_halfpt}"/><w:szCs w:val="{fs_halfpt}"/></w:rPr>'
        f'<w:t>{text}</w:t>'
        '</w:r>'
        '</w:p>'
    )


def cell_para(fs_halfpt: int, *, rule: str = None, line_tw: int = None,
              text: str = "あいうえお", font_ea: str = None, font_ascii: str = None) -> str:
    rfonts = ''
    if font_ea or font_ascii:
        ea = f' w:eastAsia="{font_ea}"' if font_ea else ''
        ac = f' w:ascii="{font_ascii}" w:hAnsi="{font_ascii}"' if font_ascii else ''
        rfonts = f'<w:rFonts{ea}{ac}/>'
    spacing = ''
    if rule and line_tw:
        spacing = f'<w:spacing w:line="{line_tw}" w:lineRule="{rule}"/>'
    return (
        '<w:p>'
        f'<w:pPr>{spacing}</w:pPr>'
        '<w:r>'
        f'<w:rPr>{rfonts}<w:sz w:val="{fs_halfpt}"/><w:szCs w:val="{fs_halfpt}"/></w:rPr>'
        f'<w:t>{text}</w:t>'
        '</w:r>'
        '</w:p>'
    )


def one_cell_table(cell_inner_xml: str, *, tcw_tw: int = 7000, trheight_tw: int = None) -> str:
    """Single-row single-cell table. Optional trHeight."""
    tr_height = ''
    if trheight_tw:
        tr_height = f'<w:trPr><w:trHeight w:val="{trheight_tw}"/></w:trPr>'
    return (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="7000" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="7000"/></w:tblGrid>'
        '<w:tr>'
        f'{tr_height}'
        '<w:tc>'
        f'<w:tcPr><w:tcW w:w="{tcw_tw}" w:type="dxa"/></w:tcPr>'
        f'{cell_inner_xml}'
        '</w:tc>'
        '</w:tr>'
        '</w:tbl>'
    )


# ============= A series — body, no table =============
A_repros = {
    # Replicate 04b88 first paragraph exactly
    "A1_body_exact340_mincho_105": body_exact(340, 21) * 4,
    # Bigger exact lh — offset hypothesis becomes more obvious
    "A2_body_exact480_mincho_105": body_exact(480, 21) * 3,
    # Replicate 1ec1 SHAPE pattern (but in body context)
    "A3_body_exact440_mincho_14":  body_exact(440, 28) * 3,
    # A1 + explicit top alignment
    "A4_body_exact340_top":        body_exact(340, 21, text_align="top") * 4,
    # A1 + explicit bottom alignment
    "A5_body_exact340_bottom":     body_exact(340, 21, text_align="bottom") * 4,
    # Single rule baseline (no exact) — Mechanism B alone (in body)
    "A6_body_single_mincho_105":   body_single(21) * 4,
    # Non-CJK font with exact
    "A7_body_exact340_TNR_105":    body_exact(340, 21).replace(
        '<w:r>',
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>',
    ).replace(
        '<w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>',
        '',
    ) * 4 if False else  # disabled — text mangling above is buggy
        body_exact(340, 21, text="Hello world.") * 4,  # use Latin text instead
}

# ============= B series — single-cell table =============
B_repros = {
    # Replicate 04b88 table cell exactly
    "B1_cell_single_mincho_105":   one_cell_table(cell_para(21)) * 2,
    # Bigger fs in cell
    "B2_cell_single_mincho_14":    one_cell_table(cell_para(28)) * 2,
    # Different CJK font
    "B3_cell_single_yumincho_105": one_cell_table(cell_para(21, font_ea="游明朝")) * 2,
    # Non-CJK font (Latin text)
    "B4_cell_single_TNR_105":      one_cell_table(cell_para(21, text="Hello world.", font_ascii="Times New Roman")) * 2,
    # Cell + exact rule
    "B5_cell_exact480_mincho_105": one_cell_table(cell_para(21, rule="exact", line_tw=480)) * 2,
    # Cell + atLeast
    "B6_cell_atLeast240_mincho_105": one_cell_table(cell_para(21, rule="atLeast", line_tw=240)) * 2,
}


if __name__ == "__main__":
    n = 0
    for name, body in {**A_repros, **B_repros}.items():
        p = make_docx(name, body_xml=body)
        n += 1
        print(f"  {p.name}")
    print(f"\nbuilt {n} repros in {OUT}")
