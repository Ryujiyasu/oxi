"""V51-V60: Isolate which feature combo with balanceSingleByteDoubleByteWidth
flag triggers Word's 18pt cell paragraph snap.

V48 = balance + jc=center + fs=18hp + vAlign center + trHeight=851 → 18pt
V49 = same minus balance → 11.5pt

Need to find: which subset of {jc=center, fs=9pt, vAlign center, trHeight=851
atLeast, multi-paragraph} when combined with balance flag fires the snap.

V51 = balance + jc=center + fs=18 + vAlign + trHeight=851 (= V48)        → control 18pt
V52 = V51 minus jc=center                                                 → ?
V53 = V51 minus fs=18 (default sz=20=10pt)                                → ?
V54 = V51 minus vAlign center                                             → ?
V55 = V51 minus trHeight                                                  → ?
V56 = V51 with 2 paragraphs (instead of 4)                                → ?
V57 = balance + jc=center only (drop fs/vAlign/trHeight)                  → ?
V58 = balance + vAlign only                                                → ?
V59 = balance + trHeight=851 only                                          → ?
V60 = V51 minus fs AND minus vAlign (drop both)                            → ?
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


def write_minimal_docx(label: str, document_xml: str):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
        zf.writestr("word/settings.xml", SETTINGS_BALANCE)
        zf.writestr("word/styles.xml", STYLES_BASIC)
        zf.writestr("word/document.xml", document_xml)
    print(f"  wrote {out_path}")


def make_doc(table_xml: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{table_xml}
<w:p/>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>"""


def make_para(label: str, line: int, *, sz_hp: int | None = None,
              jc: str | None = None) -> str:
    pPr_inner = f'<w:jc w:val="{jc}"/>' if jc else ''
    rPr = f'<w:rPr><w:sz w:val="{sz_hp}"/></w:rPr>' if sz_hp else ''
    return f'<w:p><w:pPr>{pPr_inner}</w:pPr><w:r>{rPr}<w:t>{label} L{line}</w:t></w:r></w:p>'


def make_cell_table(label: str, n_paras: int, *, sz_hp: int | None = None,
                    jc: str | None = None, valign: str | None = None,
                    trheight: int | None = None,
                    trheight_rule: str = "atLeast") -> str:
    paras = "".join(make_para(label, i, sz_hp=sz_hp, jc=jc) for i in range(1, n_paras + 1))
    valign_xml = f'<w:vAlign w:val="{valign}"/>' if valign else ''
    trh_xml = ""
    if trheight is not None:
        trh_xml = f'<w:trPr><w:trHeight w:val="{trheight}" w:hRule="{trheight_rule}"/></w:trPr>'
    return f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>{trh_xml}
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/>{valign_xml}</w:tcPr>
{paras}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT_DIR, exist_ok=True)

    # ===== V51: control reproducing V48 (= 18pt expected) =====
    write_minimal_docx("b5f706_V51_control",
                       make_doc(make_cell_table("V51", 4, sz_hp=18, jc="center",
                                                  valign="center", trheight=851)))

    # ===== V52: V51 minus jc=center =====
    write_minimal_docx("b5f706_V52_no_jc",
                       make_doc(make_cell_table("V52", 4, sz_hp=18, jc=None,
                                                  valign="center", trheight=851)))

    # ===== V53: V51 minus fs=18 (= default 10pt) =====
    write_minimal_docx("b5f706_V53_no_fs",
                       make_doc(make_cell_table("V53", 4, sz_hp=None, jc="center",
                                                  valign="center", trheight=851)))

    # ===== V54: V51 minus vAlign center =====
    write_minimal_docx("b5f706_V54_no_valign",
                       make_doc(make_cell_table("V54", 4, sz_hp=18, jc="center",
                                                  valign=None, trheight=851)))

    # ===== V55: V51 minus trHeight =====
    write_minimal_docx("b5f706_V55_no_trh",
                       make_doc(make_cell_table("V55", 4, sz_hp=18, jc="center",
                                                  valign="center", trheight=None)))

    # ===== V56: V51 with 2 paragraphs =====
    write_minimal_docx("b5f706_V56_2paras",
                       make_doc(make_cell_table("V56", 2, sz_hp=18, jc="center",
                                                  valign="center", trheight=851)))

    # ===== V57: balance + jc=center only (drop fs/vAlign/trHeight) =====
    write_minimal_docx("b5f706_V57_jc_only",
                       make_doc(make_cell_table("V57", 4, sz_hp=None, jc="center",
                                                  valign=None, trheight=None)))

    # ===== V58: balance + vAlign only =====
    write_minimal_docx("b5f706_V58_valign_only",
                       make_doc(make_cell_table("V58", 4, sz_hp=None, jc=None,
                                                  valign="center", trheight=None)))

    # ===== V59: balance + trHeight=851 only =====
    write_minimal_docx("b5f706_V59_trh_only",
                       make_doc(make_cell_table("V59", 4, sz_hp=None, jc=None,
                                                  valign=None, trheight=851)))

    # ===== V60: V51 minus fs AND minus vAlign =====
    write_minimal_docx("b5f706_V60_no_fs_no_valign",
                       make_doc(make_cell_table("V60", 4, sz_hp=None, jc="center",
                                                  valign=None, trheight=851)))


if __name__ == "__main__":
    main()
