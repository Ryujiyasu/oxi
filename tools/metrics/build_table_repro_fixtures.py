"""Author minimal `<w:tbl>` table repro fixtures.

Table / TableRow / TableCell end-to-end coverage (S294): parser at
parser/ooxml.rs:759 handles `<w:tbl>` body element. Per-row at :4572
(`<w:tr>`), per-cell at :5065 (`<w:tc>`). Cell properties:
gridSpan (h-merge), vMerge (v-merge), tcW (width), shd (shading)
parsed at :5278/:5286/:5270/:5295.

Outputs to ``tools/fixtures/table_samples/`` (committed, S272 no-COM
direct-write).

Fixtures (4):
  v1_simple_2x2.docx       — 2×2 table with text cells [A1, A2 / B1, B2]
  v1_horizontal_merge.docx — row 1 cell spans 2 columns (gridSpan=2)
  v1_vertical_merge.docx   — col 1 spans 2 rows (vMerge=restart/continue)
  v1_cell_shading.docx     — cell with shd fill="FFFF00" (yellow)
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "table_samples")

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
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="ＭＳ 明朝" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
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
    '</w:sectPr>'
)


def _cell(text: str, *, width_dxa: int = 2500, grid_span: int | None = None,
          v_merge: str | None = None, shading_fill: str | None = None) -> str:
    """Build a `<w:tc>` element with optional grid_span / v_merge / shading.

    grid_span: integer N means horizontal merge of N columns.
    v_merge: "restart" starts a vertical merge; "continue" continues one;
             None = no v_merge.
    shading_fill: hex color (e.g., "FFFF00") for cell background.
    """
    grid_span_xml = f'<w:gridSpan w:val="{grid_span}"/>' if grid_span and grid_span > 1 else ''
    v_merge_xml = ''
    if v_merge == "restart":
        v_merge_xml = '<w:vMerge w:val="restart"/>'
    elif v_merge == "continue":
        v_merge_xml = '<w:vMerge/>'
    shading_xml = f'<w:shd w:val="clear" w:color="auto" w:fill="{shading_fill}"/>' if shading_fill else ''
    return (
        '<w:tc>'
        '<w:tcPr>'
        f'<w:tcW w:w="{width_dxa}" w:type="dxa"/>'
        + grid_span_xml + v_merge_xml + shading_xml +
        '</w:tcPr>'
        f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
        '</w:tc>'
    )


def _row(*cells_xml: str) -> str:
    return f'<w:tr>{"".join(cells_xml)}</w:tr>'


def _table(grid_cols: list[int], *rows_xml: str) -> str:
    grid_xml = "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in grid_cols) + "</w:tblGrid>"
    return (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="5000" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        + grid_xml + "".join(rows_xml) +
        '</w:tbl>'
    )


def write_docx(path: str, body_xml: str) -> None:
    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_simple_2x2: 2 cols × 2 rows, text [A1, A2 / B1, B2]
    body = _table(
        [2500, 2500],
        _row(_cell("A1"), _cell("A2")),
        _row(_cell("B1"), _cell("B2")),
    )
    write_docx(os.path.join(OUT_DIR, "v1_simple_2x2.docx"), body)

    # v1_horizontal_merge: row 1 has one cell spanning both columns
    # row 2 has two regular cells
    body = _table(
        [2500, 2500],
        _row(_cell("Wide header", width_dxa=5000, grid_span=2)),
        _row(_cell("Left"), _cell("Right")),
    )
    write_docx(os.path.join(OUT_DIR, "v1_horizontal_merge.docx"), body)

    # v1_vertical_merge: column 1 spans 2 rows
    # row 1: [v_merge=restart "Tall left"] [normal "Top right"]
    # row 2: [v_merge=continue ""]          [normal "Bottom right"]
    body = _table(
        [2500, 2500],
        _row(_cell("Tall left", v_merge="restart"), _cell("Top right")),
        _row(_cell("", v_merge="continue"), _cell("Bottom right")),
    )
    write_docx(os.path.join(OUT_DIR, "v1_vertical_merge.docx"), body)

    # v1_cell_shading: 2x1 table where one cell has yellow shading
    body = _table(
        [2500, 2500],
        _row(_cell("Plain"), _cell("Yellow", shading_fill="FFFF00")),
    )
    write_docx(os.path.join(OUT_DIR, "v1_cell_shading.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
