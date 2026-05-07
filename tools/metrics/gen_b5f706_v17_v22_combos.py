"""V17-V22 minimal repros: combo + b5f706 Table 2 row 3 mimic.

Day 28 finding from Word DML: b5f706 i=96-98 cell paragraph dy = 18pt SNAP.
This cell has: 4 paragraphs, jc=center, fs=18hp(=9pt), trHeight=851tw atLeast,
inside a 3-row table with row 3 having 15 cells and other rows different.

V12 (single cell × 3 para, no jc, no trHeight, no font sz) → 13pt natural.

Hypothesis tree for trigger:
- V17: 1 cell × 3 paragraphs + trHeight=851 atLeast (no jc, default fs=20)
- V18: V17 + jc=center
- V19: V17 + fs=18 (= 9pt, like b5f706 row 3)
- V20: V17 + jc=center + fs=18 (= V18+V19 combo, full b5f706 row3 cell mimic)
- V21: V20 with trHeight=400 (smaller atLeast)
- V22: V20 with NO trHeight (= V12 + jc + fs=18)

Each tests one combination to identify which combo triggers Word's 18pt snap.
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

STYLES_BASIC = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""


def write_docx(label: str, document_xml: str):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
        zf.writestr("word/settings.xml", SETTINGS_PLAIN)
        zf.writestr("word/styles.xml", STYLES_BASIC)
        zf.writestr("word/document.xml", document_xml)
    print(f"  wrote {out_path}")


def make_doc(body: str, grid_pitch: int = 360) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>
<w:docGrid w:type="lines" w:linePitch="{grid_pitch}"/>
</w:sectPr>
</w:body>
</w:document>"""


def make_para_v(label: str, line: int, *, sz_hp: int = 20, jc: str | None = None) -> str:
    jc_xml = f'<w:jc w:val="{jc}"/>' if jc else ''
    return (
        f'<w:p><w:pPr>{jc_xml}</w:pPr>'
        f'<w:r><w:rPr><w:sz w:val="{sz_hp}"/></w:rPr>'
        f'<w:t>{label} line {line}</w:t></w:r></w:p>'
    )


def make_table(label: str, n_paras: int, *, sz_hp: int = 20, jc: str | None = None,
               trheight: int | None = None, trheight_rule: str = "atLeast") -> str:
    """Build single-row, single-cell table with N paragraphs."""
    paras = "".join(make_para_v(label, i, sz_hp=sz_hp, jc=jc) for i in range(1, n_paras + 1))
    trheight_xml = ""
    if trheight is not None:
        trheight_xml = f'<w:trPr><w:trHeight w:val="{trheight}" w:hRule="{trheight_rule}"/></w:trPr>'
    return f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>{trheight_xml}
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{paras}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT_DIR, exist_ok=True)

    # V17: 3 paragraphs + trHeight=851 atLeast (only)
    write_docx("b5f706_V17_trh851",
               make_doc(make_table("V17", 3, trheight=851)))

    # V18: V17 + jc=center
    write_docx("b5f706_V18_trh851_jc",
               make_doc(make_table("V18", 3, jc="center", trheight=851)))

    # V19: V17 + fs=9pt (sz=18)
    write_docx("b5f706_V19_trh851_fs9",
               make_doc(make_table("V19", 3, sz_hp=18, trheight=851)))

    # V20: V17 + jc=center + fs=9pt (full b5f706 row 3 cell mimic)
    write_docx("b5f706_V20_full",
               make_doc(make_table("V20", 3, sz_hp=18, jc="center", trheight=851)))

    # V21: V20 with smaller trHeight=400 atLeast
    write_docx("b5f706_V21_trh400_full",
               make_doc(make_table("V21", 3, sz_hp=18, jc="center", trheight=400)))

    # V22: V20 with NO trHeight
    write_docx("b5f706_V22_notrh_full",
               make_doc(make_table("V22", 3, sz_hp=18, jc="center", trheight=None)))


if __name__ == "__main__":
    main()
