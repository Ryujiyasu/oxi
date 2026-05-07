"""Generate V7-V16 minimal repros to test 10 hypotheses for Word's hidden
cell snap rule trigger.

Base: V4 mimic = MS Gothic 10pt (sz=20) + linePitch=360 (=18pt) + portrait + snap=1.
V4 has Word ΔY = 13pt natural (no snap). Each V7-V16 adds ONE feature from
b5f706 to test if it flips Word to dy=18pt snap.

Hypotheses (from `docs/spec/phase_beta_step3_v7_plan.md`):
- V7:  + balanceSingleByteDoubleByteWidth flag in settings.xml
- V8:  + tblStyle reference (Table Grid style with borders)
- V9:  + vAlign center on cells
- V10: + tblBorders all sides (insideH, insideV, top/bot/left/right)
- V11: + 2 cells per row (instead of 1)
- V12: + 2 rows (instead of 1)
- V13: + tcMar non-default (108tw L/R, 0 T/B)
- V14: + cell shading (w:shd val="DAE9F7")
- V15: + trHeight=400 atLeast
- V16: + mixed font sizes (9pt + 10pt + 10.5pt within cell)

Output: tools/golden-test/repros/grid_snap/b5f706_V{7..16}.docx
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

# Plain styles (V4 base = no tblStyle)
STYLES_BASIC = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

# Styles with Table Grid style (V8)
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


def write_docx(label: str, document_xml: str, settings_xml: str, styles_xml: str):
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


def build_doc(body: str, page_landscape: bool = False, grid_pitch: int = 360) -> str:
    pg_sz = ('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' if page_landscape
             else '<w:pgSz w:w="11906" w:h="16838"/>')
    grid = f'<w:docGrid w:type="lines" w:linePitch="{grid_pitch}"/>'
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr>
{pg_sz}
<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>
{grid}
</w:sectPr>
</w:body>
</w:document>"""


def make_para(label: str, line: int, sz_hp: int = 20, snap: bool = True) -> str:
    snap_xml = "" if snap else '<w:snapToGrid w:val="0"/>'
    return (
        f'<w:p><w:pPr>{snap_xml}</w:pPr>'
        f'<w:r><w:rPr><w:sz w:val="{sz_hp}"/></w:rPr>'
        f'<w:t>{label} line {line}</w:t></w:r></w:p>'
    )


def make_paras(label: str, count: int = 6, sz_hp: int = 20, snap: bool = True) -> str:
    return "".join(make_para(label, i, sz_hp, snap) for i in range(1, count + 1))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT_DIR, exist_ok=True)

    # === V7: balance flag in settings.xml ===
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{make_paras("V7", 6)}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V7_balance", build_doc(body), SETTINGS_BALANCE, STYLES_BASIC)

    # === V8: tblStyle reference (Table Grid) ===
    body = f"""<w:tbl>
<w:tblPr><w:tblStyle w:val="aa"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{make_paras("V8", 6)}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V8_tblstyle", build_doc(body), SETTINGS_PLAIN, STYLES_TBLGRID)

    # === V9: vAlign center on cell ===
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
{make_paras("V9", 6)}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V9_valign", build_doc(body), SETTINGS_PLAIN, STYLES_BASIC)

    # === V10: tblBorders all sides (inline, no tblStyle) ===
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{make_paras("V10", 6)}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V10_borders", build_doc(body), SETTINGS_PLAIN, STYLES_BASIC)

    # === V11: 2 cells per row ===
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
{make_paras("V11a", 6)}
</w:tc>
<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
{make_paras("V11b", 6)}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V11_2cells", build_doc(body), SETTINGS_PLAIN, STYLES_BASIC)

    # === V12: 2 rows ===
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{make_paras("V12r1", 3)}
</w:tc>
</w:tr>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{make_paras("V12r2", 3)}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V12_2rows", build_doc(body), SETTINGS_PLAIN, STYLES_BASIC)

    # === V13: tcMar non-default (108tw L/R, 0 T/B) ===
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>
<w:tblCellMar>
<w:top w:w="0" w:type="dxa"/>
<w:left w:w="108" w:type="dxa"/>
<w:bottom w:w="0" w:type="dxa"/>
<w:right w:w="108" w:type="dxa"/>
</w:tblCellMar>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{make_paras("V13", 6)}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V13_tcmar", build_doc(body), SETTINGS_PLAIN, STYLES_BASIC)

    # === V14: cell shading (w:shd val="DAE9F7" = b5f706 first cell color) ===
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="DAE9F7"/></w:tcPr>
{make_paras("V14", 6)}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V14_shading", build_doc(body), SETTINGS_PLAIN, STYLES_BASIC)

    # === V15: trHeight=400 atLeast (= 20pt minimum row) ===
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:trPr><w:trHeight w:val="400" w:hRule="atLeast"/></w:trPr>
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{make_paras("V15", 6)}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V15_trheight", build_doc(body), SETTINGS_PLAIN, STYLES_BASIC)

    # === V16: mixed font sizes within cell (9pt, 10pt, 10.5pt) ===
    paras_mixed = ""
    for i, sz in enumerate([18, 20, 21, 18, 20, 21], start=1):
        paras_mixed += (
            f'<w:p><w:pPr></w:pPr>'
            f'<w:r><w:rPr><w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t>V16 line {i} (sz={sz/2:g}pt)</w:t></w:r></w:p>'
        )
    body = f"""<w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>
{paras_mixed}
</w:tc>
</w:tr>
</w:tbl>
<w:p/>"""
    write_docx("b5f706_V16_mixed_fs", build_doc(body), SETTINGS_PLAIN, STYLES_BASIC)


if __name__ == "__main__":
    main()
