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

    print(f"\nGenerated {len(L_DOCS) + 2} L-category docs in {OUT_DIR}")
    print("Next: COM-measure each with measure_grid_snap_repros.py")


if __name__ == "__main__":
    main()
