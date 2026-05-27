"""Author minimal `<w:sectPr>` repro fixtures for S314 — section
properties deepening pass.

`section_integration.rs` (S290) already covers pgSz (incl. landscape
swap), pgMar (basic 4 sides), and cols (num=2). `columns_integration.rs`
(S307) deepens the cols branches. This file fills the REMAINING sectPr
surface that no integration test pinned:

  - pgBorders: 4-side PageBorders. Same parser idioms as tblBorders
    (S311): color="auto" SUPPRESSES storage (BorderDef.color stays
    None — OPPOSITE of tcBorders S310 where auto materializes to
    "000000"). val="none"/"nil" filter (line 5508): style=none → not
    stored even with sz>0. Width=0 filter: sz=0 → not stored even with
    valid style. THREE independent filters.
  - pgMar ASYMMETRIC rounding (COM-confirmed 0e7a on margin_fix
    2026-04-13):
      top    → ROUNDED to 10tw (0.5pt grid)
      bottom → EXACT twips (no rounding)
      left   → EXACT twips
      right  → EXACT twips
      header → EXACT twips
      footer → EXACT twips
    A regression that uniformly rounded or uniformly skipped rounding
    would silently shift page-break Y limits on every doc.
  - pgMar gutter ADDITIVE to margin.left (parser/ooxml.rs:5664-5666).
    NOT a separate field; the gutter is folded into left margin at
    parse time. A regression that stored gutter as separate would
    silently double-count the offset.
  - docGrid with type="lines" + linePitch → grid_line_pitch populates.
  - docGrid with linePitch BUT NO type attribute → doc_grid_no_type
    flag flips to true; grid_line_pitch stays None. NON-OBVIOUS branch
    (parser/ooxml.rs:5695-5698): docGrid being present with linePitch
    is NOT enough — type must also be set to "lines" or "linesAndChars".
    Per CLAUDE.md: "doc_grid_no_type" gates whether CJK 83/64 multiplier
    is applied (no_type=true → multiplier skipped, COM Single heights
    used).
  - pgNumType: fmt → page_number_format (verbatim string),
    start → page_number_start (u32). Both Option fields.

Outputs to ``tools/fixtures/section_properties_deepening_samples/``.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(
    os.path.dirname(__file__), "..", "fixtures",
    "section_properties_deepening_samples",
)

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
<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>
"""


def _doc_with_sectpr(sectpr_xml: str) -> str:
    return DOC_HEAD + f'<w:sectPr>{sectpr_xml}</w:sectPr>\n</w:body>\n</w:document>'


def write_docx(path: str, body_xml: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", body_xml)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


# Standard A4 portrait pgSz + simple margins helper.
PGSZ = '<w:pgSz w:w="11906" w:h="16838"/>'


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_sect_pg_borders: 4-side PageBorders + three independent
    # storage filters (val=none, sz=0, color=auto suppression).
    #   - top: style=single sz=24 color=000000 → stored (3.0pt width,
    #     color="000000")
    #   - bottom: style=single sz=24 color="auto" → stored but color
    #     SUPPRESSED (BorderDef.color stays None — OPPOSITE of
    #     tcBorders S310 where "auto" materializes to "000000")
    #   - left:  style=none → NOT stored (filter: style=none → skip)
    #   - right: style=single sz=0 → NOT stored (filter: sz>0 required)
    sectpr = (
        PGSZ +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:pgBorders w:offsetFrom="page">'
        '<w:top w:val="single" w:sz="24" w:space="24" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="24" w:space="24" w:color="auto"/>'
        '<w:left w:val="none" w:sz="24" w:space="24"/>'
        '<w:right w:val="single" w:sz="0" w:space="24" w:color="FF0000"/>'
        '</w:pgBorders>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_sect_pg_borders.docx"),
               _doc_with_sectpr(sectpr))

    # v1_sect_pgmar_asymmetric: ASYMMETRIC rounding (COM-confirmed
    # 0e7a 2026-04-13). top=1133 must round to 10tw (1133 → 1130 → 56.5pt)
    # while bottom/left/right/header/footer stay exact:
    #   - top    w=1133  → ROUND10  → 1130tw / 20 = 56.5pt
    #   - bottom w=1133  → EXACT    → 1133tw / 20 = 56.65pt
    #   - left   w=1077  → EXACT    → 1077tw / 20 = 53.85pt
    #   - right  w=1077  → EXACT    → 53.85pt
    #   - header w=851   → EXACT    → 42.55pt
    #   - footer w=992   → EXACT    → 49.6pt
    sectpr = (
        PGSZ +
        '<w:pgMar w:top="1133" w:right="1077" w:bottom="1133" w:left="1077" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_sect_pgmar_asymmetric.docx"),
               _doc_with_sectpr(sectpr))

    # v1_sect_gutter: gutter is ADDITIVE to margin.left at parse time.
    # left=1440 + gutter=720 → margin.left = (1440+720)/20 = 108pt
    # (NOT 72pt + separate gutter field).
    sectpr = (
        PGSZ +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="720"/>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_sect_gutter.docx"),
               _doc_with_sectpr(sectpr))

    # v1_sect_docgrid_lines_pitch: docGrid type="lines" linePitch=350
    # → grid_line_pitch = 350 / 20 = 17.5pt.
    sectpr = (
        PGSZ +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:docGrid w:type="lines" w:linePitch="350"/>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_sect_docgrid_lines_pitch.docx"),
               _doc_with_sectpr(sectpr))

    # v1_sect_docgrid_linesAndChars_neg_charspace (S339, 2026-05-27):
    # docGrid type="linesAndChars" + negative charSpace = compression
    # mode. Parser populates THREE fields: grid_line_pitch (linePitch/20),
    # grid_char_pitch (default_fs + charSpace/4096), grid_char_space_raw
    # (preserves raw charSpace for layout post-process). Without this
    # fixture, the linesAndChars + charSpace branch at parser/ooxml.rs:
    # 5697-5722 has NO integration coverage — would silently regress.
    #
    # b35123-realistic values: linePitch=350 (17.5pt), charSpace=-2714
    # (~-0.663pt char compression). Per S339 corpus survey, only 2/55
    # baseline docs use charSpace<0 with linesAndChars (b35123, 191cb5);
    # 6 use charSpace>0 (29dc6e family at +1453).
    sectpr = (
        PGSZ +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_sect_docgrid_linesAndChars_neg_charspace.docx"),
               _doc_with_sectpr(sectpr))

    # v1_sect_docgrid_no_type: docGrid with linePitch but NO type
    # attribute. parser/ooxml.rs:5695-5698: `grid_type.is_empty() &&
    # line_pitch > 0` → doc_grid_no_type=true, grid_line_pitch stays
    # None. Per CLAUDE.md: doc_grid_no_type gates the CJK 83/64
    # multiplier (no_type=true → multiplier SKIPPED, COM Single
    # heights used instead).
    sectpr = (
        PGSZ +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:docGrid w:linePitch="350"/>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_sect_docgrid_no_type.docx"),
               _doc_with_sectpr(sectpr))

    # v1_sect_pgnumtype: pgNumType fmt + start populate
    # page_number_format / page_number_start. Both Option fields
    # (default None when pgNumType absent).
    sectpr = (
        PGSZ +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:pgNumType w:fmt="lowerRoman" w:start="5"/>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_sect_pgnumtype.docx"),
               _doc_with_sectpr(sectpr))

    print("Done.")


if __name__ == "__main__":
    main()
