"""Author minimal w:textDirection=tbRlV repro fixtures for COM measurement.

Phase C #4 (vertical writing) campaign Round 24 (2026-04-28).
Greenfield Path B profile (only 4/184 baseline docs use textDirection,
all tcPr-level tbRlV for cell labels).

VV1: simple 1-cell table with tbRlV text, multiple font/size/length variants.
Goal: characterize how Word renders tbRlV in cell — per-char X/Y, line stacking,
column wrapping, glyph rotation pivot point.

Generated fixtures (under pipeline_data/docx/):
  VW_V1_basic.docx           — single cell, MS Mincho 10.5pt, "申請者" tbRlV
  VW_V1_long.docx            — single cell, 連絡担当窓口 (6 chars) tbRlV
  VW_V1_msmincho_14pt.docx   — same as basic but 14pt (size variant)
  VW_V1_yu_mincho.docx       — same as basic but Yu Mincho (font variant)
  VW_V1_two_cols.docx        — text long enough to wrap to 2 columns tbRlV
"""
import os
import sys
import zipfile
from xml.sax.saxutils import escape

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = "pipeline_data/docx"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault>
<w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>
<w:sz w:val="21"/>
<w:szCs w:val="21"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr>
</w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
</w:styles>"""

DOC_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
"""

SECT_PR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '<w:cols w:space="425"/>'
    '<w:docGrid w:linePitch="360"/>'
    '</w:sectPr>'
)


def _vertical_cell(text: str, cell_w_dxa: int, font_name: str = "ＭＳ 明朝", sz_halfpt: int = 21) -> str:
    """Build a single table cell with tbRlV text direction."""
    return (
        f'<w:tc>'
        f'<w:tcPr>'
        f'<w:tcW w:w="{cell_w_dxa}" w:type="dxa"/>'
        f'<w:tcBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'</w:tcBorders>'
        f'<w:textDirection w:val="tbRlV"/>'
        f'</w:tcPr>'
        f'<w:p>'
        f'<w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="{font_name}" w:eastAsia="{font_name}" w:hAnsi="{font_name}"/>'
        f'<w:sz w:val="{sz_halfpt}"/>'
        f'</w:rPr>'
        f'<w:t xml:space="preserve">{escape(text)}</w:t>'
        f'</w:r>'
        f'</w:p>'
        f'</w:tc>'
    )


def _filler_cell(cell_w_dxa: int, label: str = "(content area)") -> str:
    """Right-side filler cell for visual reference."""
    return (
        f'<w:tc>'
        f'<w:tcPr>'
        f'<w:tcW w:w="{cell_w_dxa}" w:type="dxa"/>'
        f'<w:tcBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'</w:tcBorders>'
        f'</w:tcPr>'
        f'<w:p><w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(label)}</w:t></w:r></w:p>'
        f'</w:tc>'
    )


def _table_row(*cells: str) -> str:
    return f'<w:tr><w:trPr><w:trHeight w:val="3000"/></w:trPr>{"".join(cells)}</w:tr>'


def _table(*rows: str) -> str:
    return (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="9000" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="900"/><w:gridCol w:w="8100"/></w:tblGrid>'
        + "".join(rows) +
        '</w:tbl>'
    )


def write_docx(path: str, body_xml: str) -> None:
    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


def main() -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    print(f"Writing fixtures to {OUT_DIR}/")

    # VV1_basic: 3-char tbRlV "申請者" in narrow cell
    body = _table(_table_row(
        _vertical_cell("申請者", cell_w_dxa=900),
        _filler_cell(8100, "(applicant data area)"),
    ))
    write_docx(os.path.join(OUT_DIR, "VW_V1_basic.docx"), body)

    # VV1_long: 6-char tbRlV "連絡担当窓口" in same narrow cell (still single column)
    body = _table(_table_row(
        _vertical_cell("連絡担当窓口", cell_w_dxa=900),
        _filler_cell(8100, "(contact info area)"),
    ))
    write_docx(os.path.join(OUT_DIR, "VW_V1_long.docx"), body)

    # VV1_msmincho_14pt: larger size
    body = _table(_table_row(
        _vertical_cell("申請者", cell_w_dxa=1100, sz_halfpt=28),
        _filler_cell(7900, "(14pt size variant)"),
    ))
    write_docx(os.path.join(OUT_DIR, "VW_V1_msmincho_14pt.docx"), body)

    # VV1_yu_mincho: font variant (Yu Mincho)
    body = _table(_table_row(
        _vertical_cell("申請者", cell_w_dxa=900, font_name="Yu Mincho"),
        _filler_cell(8100, "(Yu Mincho variant)"),
    ))
    write_docx(os.path.join(OUT_DIR, "VW_V1_yu_mincho.docx"), body)

    # VV1_two_cols: long text that should wrap to 2nd column
    # 30 chars at 10.5pt = 315pt of vertical extent. Cell trHeight = 3000dxa = 150pt.
    # Should force wrap to 2nd column (right-to-left line direction).
    long_text = "現在登録されている連絡担当窓口情報の更新の有無について確認"
    body = _table(_table_row(
        _vertical_cell(long_text, cell_w_dxa=1500),
        _filler_cell(7500, "(2-column wrap test)"),
    ))
    write_docx(os.path.join(OUT_DIR, "VW_V1_two_cols.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
