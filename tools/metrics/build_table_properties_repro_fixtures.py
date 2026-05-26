"""Author minimal `<w:tbl><w:tblPr>...</w:tblPr></w:tbl>` repro fixtures
for S311 table-properties coverage deepening.

`table_integration.rs` (S294) covers row/cell structure (gridSpan,
vMerge, shd). `cell_properties_integration.rs` (S310) covers tcPr.
But the `<w:tblPr>` surface (TableStyle) is not pinned by any
integration test — only fragments of it via the unit-test layer.

This file covers:
  - tblW with type="dxa" (twips/20 → 250pt) AND type="pct"
    (val/50 → percentage). The two conversions are DIFFERENT divisors
    and a regression that used the dxa divisor for pct (or vice versa)
    would silently produce wildly wrong widths.
  - tblBorders: has_inside_h flag specifically set ONLY when insideH
    is present and non-suppressed (NOT for top/bottom/left/right).
    color="auto" on a tbl border → border_color stays None
    (SUPPRESSION — OPPOSITE of cell border which materializes "auto"
    to "000000"). val="none" suppresses the border=true flag.
  - tblLayout type="fixed" → layout = Some("fixed").
  - tblInd (twips/20 → indent) + tblCellSpacing (twips/20 →
    cell_spacing) + jc → alignment.
  - tblLook attr form: firstRow/lastRow/firstColumn/lastColumn map
    directly to bool flags; noHBand/noVBand are INVERTED into
    banded_rows/banded_columns (noHBand="1" → banded_rows=false).
  - tblpPr: floating table position. tblpX/Y in twips → pt;
    horzAnchor/vertAnchor stored as strings; leftFromText etc. in
    twips → pt.

Outputs to ``tools/fixtures/table_properties_samples/``.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "table_properties_samples")

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


def _table(tblpr_xml: str, grid_widths: list, rows: list) -> str:
    grid_xml = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid_widths)
    rows_xml = ""
    for row_cells in rows:
        cells_xml = ""
        for text in row_cells:
            para = (
                f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}'
                f'</w:t></w:r></w:p>'
            )
            cells_xml += f'<w:tc><w:tcPr/>{para}</w:tc>'
        rows_xml += f'<w:tr>{cells_xml}</w:tr>'
    return (
        f'<w:tbl><w:tblPr>{tblpr_xml}</w:tblPr>'
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

    # v1_tblw_dxa_and_pct: TWO tables in the same doc.
    #   table[0]: tblW w=5000 type="dxa" → width=250pt (twips / 20).
    #   table[1]: tblW w=2500 type="pct" → width=50.0 (val / 50,
    #     50ths of a percent). DIFFERENT divisor from dxa — a
    #     regression that used /20 for pct would silently produce
    #     width=125 instead of 50.
    t0 = _table(
        '<w:tblW w:w="5000" w:type="dxa"/>',
        [2500, 2500],
        [["d0", "d1"]],
    )
    para_between = (
        '<w:p><w:r><w:t xml:space="preserve">spacer</w:t></w:r></w:p>'
    )
    t1 = _table(
        '<w:tblW w:w="2500" w:type="pct"/>',
        [2500, 2500],
        [["p0", "p1"]],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tblw_dxa_and_pct.docx"),
               t0 + para_between + t1)

    # v1_tbl_borders_inside_h: tblBorders with all 6 sides:
    #   top/bottom/left/right + insideH + insideV. has_inside_h is the
    #   side-specific flag — must be TRUE only when insideH is present
    #   and non-suppressed.
    #   - top:    val="single" sz=8 color="000000" → border=true,
    #     border_color stays "000000", border_width=1.0pt
    #   - bottom: val="single" color="auto" → SUPPRESSION on color
    #     side (border_color stays None — OPPOSITE of cell borders
    #     where color="auto" materializes to "000000").
    #   - left/right: val="single"
    #   - insideH: val="single" → has_inside_h=true
    #   - insideV: val="none" → does NOT trip border=true alone
    #     and does NOT set any has_inside_v flag (there's no such
    #     field; insideH is the only side-specific flag).
    body = _table(
        (
            '<w:tblW w:w="5000" w:type="dxa"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="8" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="8" w:color="auto"/>'
            '<w:left w:val="single" w:sz="8"/>'
            '<w:right w:val="single" w:sz="8"/>'
            '<w:insideH w:val="single" w:sz="8"/>'
            '<w:insideV w:val="none"/>'
            '</w:tblBorders>'
        ),
        [2500, 2500],
        [["a", "b"], ["c", "d"]],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tbl_borders_inside_h.docx"), body)

    # v1_tbl_layout_jc_indent: pins
    #   - tblLayout type="fixed" → layout = Some("fixed")
    #   - jc val="center" → alignment = Some("center")
    #   - tblInd w=720 → indent=36pt (twips/20)
    #   - tblCellSpacing w=100 → cell_spacing=5pt (twips/20)
    body = _table(
        (
            '<w:tblW w:w="0" w:type="auto"/>'
            '<w:tblLayout w:type="fixed"/>'
            '<w:jc w:val="center"/>'
            '<w:tblInd w:w="720" w:type="dxa"/>'
            '<w:tblCellSpacing w:w="100" w:type="dxa"/>'
        ),
        [2000, 2000],
        [["c", "d"]],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tbl_layout_jc_indent.docx"), body)

    # v1_tbl_look_attr_form: pins per-attribute tblLook with noHBand
    # inversion:
    #   firstRow=1, lastRow=0, firstColumn=1, lastColumn=0
    #   noHBand=1 → banded_rows=false (INVERTED)
    #   noVBand=0 → banded_columns=true (INVERTED)
    # A regression that dropped the noHBand inversion would silently
    # invert the visual banding of all docs.
    body = _table(
        (
            '<w:tblW w:w="0" w:type="auto"/>'
            '<w:tblLook w:firstRow="1" w:lastRow="0" '
            'w:firstColumn="1" w:lastColumn="0" '
            'w:noHBand="1" w:noVBand="0"/>'
        ),
        [2500, 2500],
        [["L", "R"]],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tbl_look_attr_form.docx"), body)

    # v1_tbl_pos_floating: pins tblpPr — the floating-table position
    # block. All four from-text distances + x/y + both anchors set.
    body = _table(
        (
            '<w:tblW w:w="3000" w:type="dxa"/>'
            '<w:tblpPr w:tblpX="1440" w:tblpY="720" '
            'w:leftFromText="180" w:rightFromText="180" '
            'w:topFromText="200" w:bottomFromText="200" '
            'w:horzAnchor="margin" w:vertAnchor="page"/>'
        ),
        [3000],
        [["F"]],
    )
    write_docx(os.path.join(OUT_DIR, "v1_tbl_pos_floating.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
