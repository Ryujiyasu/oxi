"""Author minimal `<w:tc><w:tcPr>...</w:tcPr></w:tc>` repro fixtures for
S310 (commit subject Session 309) cell-properties coverage deepening.

`table_integration.rs` (S294) already covers gridSpan, vMerge, and shd
hex; `vertical_integration.rs` covers `textDirection`. But the breadth
of `parse_cell_properties` at parser/ooxml.rs:5246 is not pinned —
specifically tcBorders (with start/end aliases), tcMar (twips→pt),
tcW (cell width twips→pt), vAlign val, and `<w:shd w:fill="auto"/>`
which SUPPRESSES storage (in contrast to an explicit hex which is
stored verbatim).

Outputs to ``tools/fixtures/cell_properties_samples/``.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "cell_properties_samples")

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


def _cell(tcpr_xml: str, text: str) -> str:
    para = (
        f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}'
        f'</w:t></w:r></w:p>'
    )
    return f'<w:tc><w:tcPr>{tcpr_xml}</w:tcPr>{para}</w:tc>'


def _table(grid_widths: list, rows: list) -> str:
    grid_xml = "".join(
        f'<w:gridCol w:w="{w}"/>' for w in grid_widths
    )
    rows_xml = "".join(
        f'<w:tr>{"".join(cells)}</w:tr>' for cells in rows
    )
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        f'<w:tblGrid>{grid_xml}</w:tblGrid>{rows_xml}</w:tbl>'
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

    # v1_tc_borders_lrtb: explicit tcBorders with all four sides. Pins:
    #   - sz=val/8: sz="8" → 1.0pt width
    #   - color="auto" → Some("000000") materialized (NOT the literal "auto")
    #   - val="none" → None (suppression; the side stays None despite
    #     the element being present)
    body = _table(
        [2500, 2500],
        [
            [
                _cell(
                    '<w:tcBorders>'
                    '<w:top w:val="single" w:sz="8" w:color="000000"/>'
                    '<w:bottom w:val="single" w:sz="16" w:color="auto"/>'
                    '<w:left w:val="dashed" w:sz="4" w:color="FF0000"/>'
                    '<w:right w:val="none" w:sz="4"/>'
                    '</w:tcBorders>',
                    "border-cell",
                ),
                _cell("", "plain-cell"),
            ]
        ],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tc_borders_lrtb.docx"), body)

    # v1_tc_borders_start_end_aliases: `<w:start>` and `<w:end>` are
    # OOXML's newer/bidi-friendly aliases. parser/ooxml.rs:5365-5366
    # routes them to `borders.left` and `borders.right` respectively.
    body = _table(
        [3000],
        [
            [
                _cell(
                    '<w:tcBorders>'
                    '<w:start w:val="single" w:sz="8" w:color="0000FF"/>'
                    '<w:end w:val="double" w:sz="12" w:color="00FF00"/>'
                    '</w:tcBorders>',
                    "start-end",
                ),
            ]
        ],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tc_borders_start_end_aliases.docx"),
               body)

    # v1_tc_margins: explicit tcMar with all four sides. Pins twips/20:
    #   - top w=100 → 5.0pt
    #   - bottom w=200 → 10.0pt
    #   - left w=300 → 15.0pt
    #   - right w=400 → 20.0pt
    # The four sides are different so a regression that mis-routes
    # them (e.g., swap top/bottom) is caught.
    body = _table(
        [4000],
        [
            [
                _cell(
                    '<w:tcMar>'
                    '<w:top w:w="100" w:type="dxa"/>'
                    '<w:bottom w:w="200" w:type="dxa"/>'
                    '<w:left w:w="300" w:type="dxa"/>'
                    '<w:right w:w="400" w:type="dxa"/>'
                    '</w:tcMar>',
                    "padded",
                ),
            ]
        ],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tc_margins.docx"), body)

    # v1_tc_width_valign: distinguish three cells:
    #   - cell[0]: tcW=3000tw → width=150pt; vAlign="center"
    #   - cell[1]: tcW=2000tw → width=100pt; vAlign="bottom"
    #   - cell[2]: no tcW, no vAlign → width=None, v_align=None
    body = _table(
        [3000, 2000, 1000],
        [
            [
                _cell(
                    '<w:tcW w:w="3000" w:type="dxa"/>'
                    '<w:vAlign w:val="center"/>',
                    "wide-center",
                ),
                _cell(
                    '<w:tcW w:w="2000" w:type="dxa"/>'
                    '<w:vAlign w:val="bottom"/>',
                    "narrow-bottom",
                ),
                _cell("", "no-overrides"),
            ]
        ],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tc_width_valign.docx"), body)

    # v1_tc_shd_auto_suppression: pins parser/ooxml.rs:5312
    #   - cell[0]: `<w:shd w:fill="auto"/>` → shading=None (SUPPRESSED,
    #     NOT stored as the string "auto")
    #   - cell[1]: `<w:shd w:fill="FF0000"/>` → shading=Some("FF0000")
    # Symmetric with the rPr color="auto"→None branch pinned in S309,
    # but on the cell-level shd field.
    body = _table(
        [2500, 2500],
        [
            [
                _cell('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>',
                      "shd-auto"),
                _cell('<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>',
                      "shd-red"),
            ]
        ],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tc_shd_auto_suppression.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
