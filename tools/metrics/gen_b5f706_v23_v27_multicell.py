"""V23-V27 minimal repros: multi-cell row + multi-paragraph cell to find the
trigger of Word's 18pt cell paragraph snap behavior in b5f706.

V20 (1 row × 1 cell × 3 paragraphs + jc=center + fs=9pt + trHeight=851) → 11.5pt
b5f706 Table 2 row 3 cell[2] (1 cell within 15-cell row × 4 paragraphs + same params) → 18pt

Difference must be in row structure (multi-cell, mixed paragraph counts, etc.).

V23: 1 row × 15 cells × 4 paragraphs each (jc=center, fs=9pt, trHeight=851)
V24: V23 with mixed paragraph counts (some cells 4, some 1)
V25: V23 + 2 prior rows with different cell counts (mimics 3-row b5f706 Table 2)
V26: 3-row mimic exact (row1=12 cells gridSpan, row2=23 cells, row3=15 cells)
V27: V25 + tblStyle Table Grid + balanceSingleByteDoubleByteWidth flag
"""
from __future__ import annotations

import os
import sys
import zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.join(REPO, "tools", "golden-test", "repros", "grid_snap")

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

SETTINGS_PLAIN = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

SETTINGS_BALANCE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:balanceSingleByteDoubleByteWidth/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

STYLES_BASIC = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

STYLES_TBLGRID = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>
<w:style w:type="table" w:styleId="aa">
<w:name w:val="Table Grid"/>
<w:basedOn w:val="TableNormal"/>
<w:tblPr>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders>
</w:tblPr>
</w:style>
</w:styles>"""


def write_docx(label: str, document_xml: str, settings_xml=SETTINGS_PLAIN,
               styles_xml=STYLES_BASIC):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
        zf.writestr("word/settings.xml", settings_xml)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/document.xml", document_xml)
    print(f"  wrote {out_path}")


def make_doc_landscape(body: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr>
<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/>
<w:titlePg/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>"""


def make_para(label: str, line: int, sz_hp: int = 18, jc: str = "center") -> str:
    return (
        f'<w:p><w:pPr><w:jc w:val="{jc}"/></w:pPr>'
        f'<w:r><w:rPr><w:sz w:val="{sz_hp}"/></w:rPr>'
        f'<w:t>{label} L{line}</w:t></w:r></w:p>'
    )


def make_cell(label: str, n_paras: int, cell_w: int, sz_hp: int = 18,
              jc: str = "center", with_tblstyle: bool = False) -> str:
    paras = "".join(make_para(label, i, sz_hp, jc) for i in range(1, n_paras + 1))
    return f'<w:tc><w:tcPr><w:tcW w:w="{cell_w}" w:type="dxa"/></w:tcPr>{paras}</w:tc>'


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT_DIR, exist_ok=True)

    # ===== V23: 1 row × 15 cells × 4 paragraphs each (uniform) =====
    cell_w = 1000  # 15 × 1000 = 15000tw ≈ 750pt
    cells_xml = "".join(make_cell(f"V23c{i}", 4, cell_w) for i in range(1, 16))
    body_v23 = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid>{"".join(f'<w:gridCol w:w="{cell_w}"/>' for _ in range(15))}</w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="851" w:hRule="atLeast"/></w:trPr>
{cells_xml}
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V23_15cells_4para", make_doc_landscape(body_v23))

    # ===== V24: 1 row × 15 cells, mixed paragraph counts (some 4, some 1) =====
    para_counts = [4, 1, 4, 1, 4, 1, 4, 1, 4, 1, 4, 1, 4, 1, 4]
    cells_xml = "".join(make_cell(f"V24c{i}", n, cell_w) for i, n in enumerate(para_counts, 1))
    body_v24 = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid>{"".join(f'<w:gridCol w:w="{cell_w}"/>' for _ in range(15))}</w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="851" w:hRule="atLeast"/></w:trPr>
{cells_xml}
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V24_mixed_paras", make_doc_landscape(body_v24))

    # ===== V25: 3 rows mimic - row1 12 cells gridSpan=2, row2 23 cells, row3 15 cells =====
    # Row 1: 12 cells, first cell w=851, rest gridSpan=2 with w=1327 each (matching b5f706)
    row1_cells = '<w:tc><w:tcPr><w:tcW w:w="851" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
    for i in range(1, 12):
        row1_cells += f'<w:tc><w:tcPr><w:tcW w:w="1327" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc>'

    # Row 2: 23 cells × 1 paragraph
    grid_widths_23 = [851] + [663, 664] * 11
    row2_cells = "".join(make_cell(f"V25r2c{i}", 1, grid_widths_23[i-1])
                          for i in range(1, 24))

    # Row 3: 15 cells, mixed paragraph counts (mimics Table 2 row 3)
    para_counts_r3 = [1, 1, 4, 1, 4, 1, 4, 1, 4, 1, 4, 1, 4, 1, 4]
    grid_widths_r3 = [851 + 663] + [664+663] * 7  # this is approximate; doesn't fit 15 - workaround
    # simpler: just use fixed widths summing to 15451 (b5f706 table width)
    sum_w = 15451
    cell_w_r3 = sum_w // 15
    row3_cells = "".join(make_cell(f"V25r3c{i}", n, cell_w_r3)
                          for i, n in enumerate(para_counts_r3, 1))

    grid_cols = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid_widths_23)
    body_v25 = f"""<w:tbl>
<w:tblPr><w:tblW w:w="15451" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid>{grid_cols}</w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="340" w:hRule="atLeast"/></w:trPr>{row1_cells}</w:tr>
<w:tr><w:trPr><w:trHeight w:val="284" w:hRule="atLeast"/></w:trPr>{row2_cells}</w:tr>
<w:tr><w:trPr><w:trHeight w:val="851" w:hRule="atLeast"/></w:trPr>{row3_cells}</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V25_3rows_mimic", make_doc_landscape(body_v25))

    # ===== V26: V25 + tblStyle Table Grid + tblBorders =====
    body_v26 = f"""<w:tbl>
<w:tblPr><w:tblStyle w:val="aa"/><w:tblW w:w="15451" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid>{grid_cols}</w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="340" w:hRule="atLeast"/></w:trPr>{row1_cells}</w:tr>
<w:tr><w:trPr><w:trHeight w:val="284" w:hRule="atLeast"/></w:trPr>{row2_cells.replace("V25r2c","V26r2c")}</w:tr>
<w:tr><w:trPr><w:trHeight w:val="851" w:hRule="atLeast"/></w:trPr>{row3_cells.replace("V25r3c","V26r3c")}</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V26_full_mimic", make_doc_landscape(body_v26),
               settings_xml=SETTINGS_BALANCE, styles_xml=STYLES_TBLGRID)

    # ===== V27: V26 + balanceSingleByteDoubleByteWidth flag (already added in V26) =====
    # Skip duplicate - V27 = V26 minus balance
    body_v27 = body_v26.replace("V26r","V27r")
    write_docx("b5f706_V27_full_no_balance", make_doc_landscape(body_v27),
               settings_xml=SETTINGS_PLAIN, styles_xml=STYLES_TBLGRID)


if __name__ == "__main__":
    main()
