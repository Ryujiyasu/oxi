"""Generate minimal repro docx files for grid-snap refactor (Day 23+).

Per docs/spec/grid_snap_refactor_plan.md, isolates each grid-snap category
(L: line height, R: row height, C: cursor advance, E: estimate) with
minimal docx that varies one parameter at a time.

Output: tools/golden-test/repros/grid_snap/{L1..L8, R1..R6, C1..C4, E1, E2}.docx

Run: python tools/metrics/gen_grid_snap_repros.py
"""
from __future__ import annotations

import os
import sys
import zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.join(REPO, "tools", "golden-test", "repros", "grid_snap")


# Common boilerplate
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

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

STYLES_BASIC = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""


def make_docx(out_path: str, body_inner: str, doc_grid: str = '<w:docGrid w:type="lines" w:linePitch="330"/>'):
    """Create a docx with the given body content."""
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body_inner}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>
{doc_grid}
</w:sectPr>
</w:body>
</w:document>"""

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
        zf.writestr("word/settings.xml", SETTINGS)
        zf.writestr("word/styles.xml", STYLES_BASIC)
        zf.writestr("word/document.xml", document_xml)


def make_para(text: str, snap: bool, line: int | None = None, line_rule: str | None = None,
              sz_hp: int = 21):
    """Build a single <w:p> with given snap and line spacing."""
    snap_xml = "" if snap else '<w:snapToGrid w:val="0"/>'
    line_xml = ""
    if line is not None and line_rule is not None:
        line_xml = f'<w:spacing w:line="{line}" w:lineRule="{line_rule}"/>'
    return f"""<w:p>
<w:pPr>{snap_xml}{line_xml}</w:pPr>
<w:r><w:rPr><w:sz w:val="{sz_hp}"/></w:rPr><w:t>{text}</w:t></w:r>
</w:p>"""


# ===== L category: Line height under various conditions =====
# Each L doc: 6 paragraphs in a row, each labeled, to measure inter-paragraph Y diff

L_DOCS = {
    # L1: snap=1, line=Single (auto), grid=lines pitch=330 → expect snap to 16.5pt
    "L1": {
        "grid_pitch": 330,
        "paras": [(f"L1 line {i}", True, None, None, 21) for i in range(1, 7)],
    },
    # L2: snap=1, line=Single (auto), grid=lines pitch=360 → expect snap to 18pt
    "L2": {
        "grid_pitch": 360,
        "paras": [(f"L2 line {i}", True, None, None, 21) for i in range(1, 7)],
    },
    # L3: snap=0, line=Single (auto), grid=lines pitch=330 → expect natural line height
    "L3": {
        "grid_pitch": 330,
        "paras": [(f"L3 line {i}", False, None, None, 21) for i in range(1, 7)],
    },
    # L4: snap=1, line=exact 240tw=12pt, grid=lines pitch=330 → expect exact 12pt
    "L4": {
        "grid_pitch": 330,
        "paras": [(f"L4 line {i}", True, 240, "exact", 21) for i in range(1, 7)],
    },
    # L5: snap=1, line=Multiple 1.15 (= 240*1.15=276), grid=lines pitch=330
    # COM-noted: Multiple does NOT snap, so expect natural × 1.15
    "L5": {
        "grid_pitch": 330,
        "paras": [(f"L5 line {i}", True, 276, "auto", 21) for i in range(1, 7)],
    },
    # L6: snap=1, line=atLeast 280tw=14pt, grid=lines pitch=330 → expect max(snap, 14)
    "L6": {
        "grid_pitch": 330,
        "paras": [(f"L6 line {i}", True, 280, "atLeast", 21) for i in range(1, 7)],
    },
    # L7/L8: in-cell variants — built separately as table
}

# In-cell L7/L8: 1 cell with 6 paragraphs
def make_cell_doc(grid_pitch: int, snap: bool, label: str):
    paras_xml = ""
    for i in range(1, 7):
        snap_xml = "" if snap else '<w:snapToGrid w:val="0"/>'
        paras_xml += f"""<w:p>
<w:pPr>{snap_xml}</w:pPr>
<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>{label} line {i}</w:t></w:r>
</w:p>"""
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:tc>
<w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{paras_xml}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    grid = f'<w:docGrid w:type="lines" w:linePitch="{grid_pitch}"/>'
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    make_docx(out_path, body, grid)


def make_table_row(cells_content: list[str], tr_height: int | None = None,
                   tr_rule: str | None = None) -> str:
    """Build a single <w:tr> with cells. cells_content is list of inner XML for <w:tc>."""
    height_xml = ""
    if tr_height is not None:
        rule_attr = f' w:hRule="{tr_rule}"' if tr_rule else ''
        height_xml = f'<w:trHeight w:val="{tr_height}"{rule_attr}/>'
    cells_xml = ""
    for content in cells_content:
        cells_xml += f"""<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{content}</w:tc>"""
    return f"""<w:tr><w:trPr>{height_xml}</w:trPr>{cells_xml}</w:tr>"""


def make_table_doc(label: str, rows_xml: str, num_cols: int = 1, grid_pitch: int = 330):
    """Wrap rows in a <w:tbl> with appropriate gridCol setup."""
    grid_cols = "".join(f'<w:gridCol w:w="9000"/>' for _ in range(num_cols))
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid>{grid_cols}</w:tblGrid>
{rows_xml}
</w:tbl>
<w:p/>"""
    grid = f'<w:docGrid w:type="lines" w:linePitch="{grid_pitch}"/>'
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    make_docx(out_path, body, grid)


def make_para_simple(text: str, snap: bool = True, sz_hp: int = 21) -> str:
    """Simple paragraph for cell content."""
    snap_xml = "" if snap else '<w:snapToGrid w:val="0"/>'
    return f"""<w:p><w:pPr>{snap_xml}</w:pPr><w:r><w:rPr><w:sz w:val="{sz_hp}"/></w:rPr><w:t>{text}</w:t></w:r></w:p>"""


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT_DIR, exist_ok=True)

    # L1-L6: body paragraphs
    for label, spec in L_DOCS.items():
        body_inner = ""
        for p in spec["paras"]:
            body_inner += make_para(*p)
        grid = f'<w:docGrid w:type="lines" w:linePitch="{spec["grid_pitch"]}"/>'
        out_path = os.path.join(OUT_DIR, f"{label}.docx")
        make_docx(out_path, body_inner, grid)
        print(f"  {label}.docx written")

    # L7: snap=1 in cell, grid=lines pitch=330
    make_cell_doc(330, True, "L7")
    print(f"  L7.docx written")

    # L8: snap=0 in cell, grid=lines pitch=330
    make_cell_doc(330, False, "L8")
    print(f"  L8.docx written")

    # R-category design: each doc has 2+ rows so row-1 height = row-2.y_top - row-1.y_top
    # Row-2 always contains a single short paragraph (reference) so its height is fixed.

    # R1: row1 = 1 cell × 1 line of 10.5pt content, row2 = reference
    row1 = make_table_row([make_para_simple("R1 content")])
    row2 = make_table_row([make_para_simple("R-ref")])
    make_table_doc("R1", row1 + row2)
    print(f"  R1.docx written")

    # R2: row1 = 1 cell × 3 lines, row2 = reference
    paras_3 = "".join(make_para_simple(f"R2 line {i}") for i in range(1, 4))
    row1 = make_table_row([paras_3])
    make_table_doc("R2", row1 + row2)
    print(f"  R2.docx written")

    # R3: row1 = 2 cells (cell1=1 line, cell2=3 lines), row2 = reference (2 cells × ref)
    cell1 = make_para_simple("R3 cell1")
    cell2 = "".join(make_para_simple(f"R3 cell2 line {i}") for i in range(1, 4))
    row1 = make_table_row([cell1, cell2])
    row2_2col = make_table_row([make_para_simple("R-ref-1"), make_para_simple("R-ref-2")])
    make_table_doc("R3", row1 + row2_2col, num_cols=2)
    print(f"  R3.docx written")

    # R4: row1 = trHeight=400tw=20pt exact, row2 = reference
    row1 = make_table_row([make_para_simple("R4 content")], tr_height=400, tr_rule="exact")
    make_table_doc("R4", row1 + row2)
    print(f"  R4.docx written")

    # R5: row1 = trHeight=400tw=20pt atLeast, row2 = reference
    row1 = make_table_row([make_para_simple("R5 content")], tr_height=400, tr_rule="atLeast")
    make_table_doc("R5", row1 + row2)
    print(f"  R5.docx written")

    # R6: row1 = nested table inside cell, row2 = reference
    inner_tbl = """<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="4500"/></w:tblGrid>
<w:tr><w:trPr/><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>""" + make_para_simple("R6 nested cell") + """</w:tc></w:tr>
</w:tbl>"""
    cell_with_nested = inner_tbl + "<w:p/>"
    row1 = make_table_row([cell_with_nested])
    make_table_doc("R6", row1 + row2)
    print(f"  R6.docx written")

    # ===== C category: cursor advance under various conditions =====
    # 6 paragraphs with specific spacing; measure cumulative y to detect
    # whether cursor sticky-anchors to grid lines or accumulates natural.

    # C1: linesAndChars (LM2) mode, pitch=330, body paragraphs snap=1
    body_inner = "".join(make_para(f"C1 line {i}", True, None, None, 21) for i in range(1, 7))
    grid = '<w:docGrid w:type="linesAndChars" w:linePitch="330" w:charSpace="0"/>'
    out_path = os.path.join(OUT_DIR, "C1.docx")
    make_docx(out_path, body_inner, grid)
    print(f"  C1.docx written")

    # C2: mixed snap=1 + snap=0 paragraphs alternating
    paras = []
    for i in range(1, 7):
        snap = (i % 2 == 1)
        paras.append(make_para(f"C2 line {i} snap={int(snap)}", snap, None, None, 21))
    body_inner = "".join(paras)
    grid = '<w:docGrid w:type="lines" w:linePitch="330"/>'
    out_path = os.path.join(OUT_DIR, "C2.docx")
    make_docx(out_path, body_inner, grid)
    print(f"  C2.docx written")

    # C3: 6 paragraphs with line=exact 13pt (260tw)
    body_inner = "".join(make_para(f"C3 line {i}", True, 260, "exact", 21) for i in range(1, 7))
    grid = '<w:docGrid w:type="lines" w:linePitch="330"/>'
    out_path = os.path.join(OUT_DIR, "C3.docx")
    make_docx(out_path, body_inner, grid)
    print(f"  C3.docx written")

    # C4: 6 paragraphs with line=Multiple 1.5 (= 360tw=18pt for 12pt font, snapped × 1.5)
    body_inner = "".join(make_para(f"C4 line {i}", True, 360, "auto", 21) for i in range(1, 7))
    grid = '<w:docGrid w:type="lines" w:linePitch="330"/>'
    out_path = os.path.join(OUT_DIR, "C4.docx")
    make_docx(out_path, body_inner, grid)
    print(f"  C4.docx written")

    # ===== E category: estimate behavior =====
    # E1: cell with multiple paragraphs (test cell line height accumulation)
    paras_5 = "".join(make_para_simple(f"E1 line {i}") for i in range(1, 6))
    row = make_table_row([paras_5])
    make_table_doc("E1", row, grid_pitch=330)
    print(f"  E1.docx written")

    # E2: body with multiple paragraphs of different sizes (mixed line heights)
    body_inner = ""
    for i, sz in enumerate([16, 21, 28, 21, 16, 21], 1):  # 8/10.5/14/10.5/8/10.5pt
        body_inner += make_para(f"E2 line {i} sz={sz//2}pt", True, None, None, sz)
    grid = '<w:docGrid w:type="lines" w:linePitch="330"/>'
    out_path = os.path.join(OUT_DIR, "E2.docx")
    make_docx(out_path, body_inner, grid)
    print(f"  E2.docx written")

    print(f"\nGenerated {len(L_DOCS) + 2 + 6 + 4 + 2} (L+R+C+E) docs in {OUT_DIR}")
    print("Next: COM-measure each with measure_grid_snap_repros.py")


if __name__ == "__main__":
    main()
