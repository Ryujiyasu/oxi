"""Day 33 part 68 (R7.21 a47e6 campaign) — Minimal repros for nested-table
atLeast row precision investigation.

Context: a47e6 table 1 row 6 has a nested 8-row table at fs=7pt with
trHeight=17 (atLeast). Day 33 part 61 measured +3.5pt cumulative over-pump
in this nested-table cell (~0.44pt/row × 8). This drives the +1.4pt
overflow at pi=2 备考 → 用紙 boundary that prevents a47e6 from being 1-page
like Word.

This generator builds variants isolating which factor drives the per-row
over-pump:

| Variant | rows | fs | hRule | val | parent cells |
|---|---|---|---|---|---|
| N01_base   | 8 | 7  | atLeast | 17 | 2 |
| N02_4row   | 4 | 7  | atLeast | 17 | 2 |
| N03_2row   | 2 | 7  | atLeast | 17 | 2 |
| N04_1row   | 1 | 7  | atLeast | 17 | 2 |
| N05_fs10   | 8 | 10 | atLeast | 17 | 2 |
| N06_fs105  | 8 | 10.5 | atLeast | 17 | 2 |
| N07_auto   | 8 | 7  | (none)  | -  | 2 |
| N08_exact  | 8 | 7  | exact   | 17 | 2 |
| N09_val24  | 8 | 7  | atLeast | 24 | 2 |
| N10_1col   | 8 | 7  | atLeast | 17 | 1 |
| N11_fs9_v20| 8 | 9  | atLeast | 20 | 2 |

Word COM measures each row's start Y; compare to Oxi BR_DUMP / TBL_DUMP.

Run: python tools/metrics/gen_nested_atleast_repro.py
Output: tools/golden-test/repros/nested_atleast/N*.docx
"""

from __future__ import annotations
import os
import zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'nested_atleast')

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

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

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:hAnsi="MS Mincho" w:eastAsia="MS Mincho"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>"""


def nested_table_xml(n_rows: int, fs_half: int, h_rule: str | None, h_val: int) -> str:
    """Build XML for the nested table.

    fs_half: half-point font size (7pt → 14, 10pt → 20)
    h_rule: "atLeast" / "exact" / None (= no trHeight)
    h_val: trHeight value in twentieths-of-a-point (= twips)
    """
    rows = []
    for r in range(n_rows):
        # 2 cells per nested row (mirrors a47e6's nested structure)
        trh_xml = ''
        if h_rule:
            trh_xml = f'<w:trHeight w:val="{h_val}" w:hRule="{h_rule}"/>'
        elif h_val:
            trh_xml = f'<w:trHeight w:val="{h_val}"/>'
        row_xml = f"""<w:tr><w:trPr>{trh_xml}</w:trPr>
<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
<w:p><w:pPr><w:rPr><w:sz w:val="{fs_half}"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="{fs_half}"/></w:rPr><w:t>R{r+1}c0</w:t></w:r></w:p>
</w:tc>
<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
<w:p><w:pPr><w:rPr><w:sz w:val="{fs_half}"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="{fs_half}"/></w:rPr><w:t>R{r+1}c1</w:t></w:r></w:p>
</w:tc>
</w:tr>"""
        rows.append(row_xml)

    return f"""<w:tbl>
<w:tblPr>
<w:tblW w:w="4000" w:type="dxa"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
{''.join(rows)}
</w:tbl>"""


def doc_xml(n_rows: int, fs_half: int, h_rule: str | None, h_val: int, parent_cols: int) -> str:
    """Build full document.xml with parent table containing nested table."""
    nested = nested_table_xml(n_rows, fs_half, h_rule, h_val)
    # Parent table has 1 row with 1 or 2 cells
    if parent_cols == 1:
        parent_row = f"""<w:tr>
<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
<w:p/>{nested}<w:p/>
</w:tc>
</w:tr>"""
        grid = '<w:gridCol w:w="4500"/>'
    else:
        parent_row = f"""<w:tr>
<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
<w:p><w:r><w:t>(empty cell)</w:t></w:r></w:p>
</w:tc>
<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
<w:p/>{nested}<w:p/>
</w:tc>
</w:tr>"""
        grid = '<w:gridCol w:w="2000"/><w:gridCol w:w="4500"/>'

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>HEADER MARKER</w:t></w:r></w:p>
<w:tbl>
<w:tblPr>
<w:tblW w:w="0" w:type="auto"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid>{grid}</w:tblGrid>
{parent_row}
</w:tbl>
<w:p><w:r><w:t>FOOTER MARKER</w:t></w:r></w:p>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>"""


def build(name: str, n_rows: int, fs_pt: float, h_rule: str | None, h_val: int, parent_cols: int = 2) -> str:
    fs_half = int(fs_pt * 2)
    out_path = os.path.join(OUT_DIR, f'{name}.docx')
    os.makedirs(OUT_DIR, exist_ok=True)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', doc_xml(n_rows, fs_half, h_rule, h_val, parent_cols))
    return out_path


VARIANTS = [
    # (name, n_rows, fs_pt, h_rule, h_val_twips, parent_cols)
    ('N01_base_8r_fs7_aL17',   8, 7,    'atLeast', 340, 2),  # 17pt = 340 twips
    ('N02_4r_fs7_aL17',        4, 7,    'atLeast', 340, 2),
    ('N03_2r_fs7_aL17',        2, 7,    'atLeast', 340, 2),
    ('N04_1r_fs7_aL17',        1, 7,    'atLeast', 340, 2),
    ('N05_8r_fs10_aL17',       8, 10,   'atLeast', 340, 2),
    ('N06_8r_fs105_aL17',      8, 10.5, 'atLeast', 340, 2),
    ('N07_8r_fs7_auto',        8, 7,    None,        0, 2),
    ('N08_8r_fs7_exact17',     8, 7,    'exact',   340, 2),
    ('N09_8r_fs7_aL24',        8, 7,    'atLeast', 480, 2),  # 24pt = 480 twips
    ('N10_8r_fs7_aL17_1col',   8, 7,    'atLeast', 340, 1),
    ('N11_8r_fs9_aL20',        8, 9,    'atLeast', 400, 2),  # 20pt = 400 twips
]


def main() -> int:
    for v in VARIANTS:
        path = build(*v)
        print(f'  built: {path}')
    print(f'\n{len(VARIANTS)} variants -> {OUT_DIR}')
    return 0


if __name__ == '__main__':
    import sys
    sys.exit(main())
