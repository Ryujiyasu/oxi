"""Author minimal `<w:tr><w:trPr>...</w:trPr></w:tr>` repro fixtures for
S312 row-properties coverage deepening.

`table_integration.rs` (S294) covers row/cell structure (gridSpan,
vMerge, shd). `cell_properties_integration.rs` (S310) covers tcPr.
`table_properties_integration.rs` (S311) covers tblPr. This file
fills the remaining `<w:trPr>` (row-level) surface that no
integration test pinned:

  - trHeight: val/20 twips→pt, AND hRule stored verbatim as the
    enum string ("exact" vs "atLeast"). These two rules drive
    radically different layout behavior (clip vs grow), so a
    regression that dropped hRule storage would silently break
    every "exact" row across the corpus.
  - tblHeader → row.header=true (repeat-as-header-on-each-page).
  - cantSplit → row.cant_split=true (forbid row-internal page break).
  - gridBefore val=2 → row.grid_before=2 (skip 2 leading grid
    columns; lets a row start at column N>0).
  - tblPrEx > tblCellMar: ROW-level cell margin override (separate
    from the table-default tblCellMar pinned in S311). The
    start/end aliases also route to left/right at the parser
    branch parser/ooxml.rs:5104-5105.

Outputs to ``tools/fixtures/row_properties_samples/``.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "row_properties_samples")

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

DOC_TAIL = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '</w:sectPr>'
    '\n</w:body>\n</w:document>'
)


def _cell(text: str) -> str:
    para = (
        f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}'
        f'</w:t></w:r></w:p>'
    )
    return f'<w:tc><w:tcPr/>{para}</w:tc>'


def _row(trpr_xml: str, cells_text: list,
         tblpr_ex_xml: str = "") -> str:
    cells_xml = "".join(_cell(t) for t in cells_text)
    # ECMA-376 spec: <w:tr> children appear in order:
    #   <w:tblPrEx>? <w:trPr>? <w:tc>+
    # tblPrEx is a SIBLING of trPr, NOT a child.
    return (
        f'<w:tr>{tblpr_ex_xml}<w:trPr>{trpr_xml}</w:trPr>{cells_xml}</w:tr>'
    )


def _table(grid_widths: list, rows_xml: list) -> str:
    grid_xml = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid_widths)
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        f'<w:tblGrid>{grid_xml}</w:tblGrid>{"".join(rows_xml)}</w:tbl>'
    )


def write_docx(path: str, body_xml: str) -> None:
    full = DOC_HEAD + body_xml + DOC_TAIL
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

    # v1_tr_height_exact: trHeight val=400 + hRule="exact". Pins:
    #   - val/20 conversion (400tw → 20pt)
    #   - hRule stored verbatim as "exact" (the layout dispatches on
    #     this string: "exact" means CLIP-to-height, no auto-grow)
    body = _table(
        [3000],
        [
            _row(
                '<w:trHeight w:val="400" w:hRule="exact"/>',
                ["exact-20pt"],
            ),
        ],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tr_height_exact.docx"), body)

    # v1_tr_height_atleast: trHeight val=600 + hRule="atLeast". The
    # "atLeast" rule lets the row GROW past 30pt if content overflows
    # — opposite policy from "exact". A regression that dropped
    # hRule would conflate the two policies → over-clip atLeast rows
    # or over-grow exact rows.
    body = _table(
        [3000],
        [
            _row(
                '<w:trHeight w:val="600" w:hRule="atLeast"/>',
                ["atLeast-30pt"],
            ),
        ],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tr_height_atleast.docx"), body)

    # v1_tr_header_cant_split: two flags on the same row.
    #   tblHeader → header=true (repeat row at top of every page)
    #   cantSplit → cant_split=true (forbid row from page-break)
    # Independent flags — both can be true on the same row.
    body = _table(
        [3000],
        [
            _row(
                '<w:tblHeader/><w:cantSplit/>',
                ["header-cant-split"],
            ),
            _row("", ["plain-row"]),
        ],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tr_header_cant_split.docx"), body)

    # v1_tr_grid_before: gridBefore val=2 means the row STARTS at
    # grid column 2 (skips 2 leading columns). Row has 1 cell that
    # visually appears at column 2; the table grid still has 3
    # columns. Layout uses this to indent the start of the row.
    body = _table(
        [1500, 1500, 1500],
        [
            _row(
                '<w:gridBefore w:val="2"/>',
                ["far-right"],
            ),
            _row("", ["col0", "col1", "col2"]),
        ],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tr_grid_before.docx"), body)

    # v1_tr_tblpr_ex_cellmar_override: tblPrEx is a ROW-level escape
    # hatch that overrides table-level tblPr for this row only.
    # Specifically tblPrEx > tblCellMar populates
    # row.cell_margins_override (NOT the table-level default which
    # tblPr > tblCellMar populates — pinned in S311). Also uses
    # `<w:start>` / `<w:end>` aliases to confirm parser/ooxml.rs:5104
    # routes them to left/right at THIS code path too.
    body = _table(
        [3000],
        [
            _row(
                "",  # trPr empty
                ["row-override"],
                tblpr_ex_xml=(
                    '<w:tblPrEx>'
                    '<w:tblCellMar>'
                    '<w:top w:w="100" w:type="dxa"/>'
                    '<w:bottom w:w="200" w:type="dxa"/>'
                    '<w:start w:w="300" w:type="dxa"/>'
                    '<w:end w:w="400" w:type="dxa"/>'
                    '</w:tblCellMar>'
                    '</w:tblPrEx>'
                ),
            ),
        ],
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_tr_tblpr_ex_cellmar_override.docx"),
        body,
    )

    print("Done.")


if __name__ == "__main__":
    main()
