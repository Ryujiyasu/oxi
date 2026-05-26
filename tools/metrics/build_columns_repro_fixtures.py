"""Author minimal sectPr/cols repro fixtures for ColumnLayout deepening (S307).

section_samples already covers num=1 default and num=2 + space. This
file ships fixtures that exercise the parts the parser at
parser/ooxml.rs:5529 / :5779 implements but section_integration.rs
doesn't yet pin:
  - num=1 in XML → Page.columns stays None (parser short-circuits
    `if num > 1` on line 5574 / 5793)
  - equalWidth="0" → ColumnLayout.equal_width = false; absent or other
    value → true (default semantics from OOXML)
  - <w:cols> as a Start element with child <w:col> entries → each
    child populates ColumnDef { width, space } from w:w / w:space
    attributes (line 5550 / 5562)
  - <w:cols/> as Empty element shape → same num/space/equalWidth
    attrs but no col_defs (line 5779)
  - num=3 with explicit space — confirms multi-column space scaling
  - Per-col individual space overrides global space

Outputs to ``tools/fixtures/columns_samples/``.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "columns_samples")

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


def _paragraph(text: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def _sect_pr(cols_xml: str) -> str:
    return (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        f'{cols_xml}'
        '</w:sectPr>'
    )


def write_docx(path: str, body_xml: str) -> None:
    full = DOC_HEAD + body_xml + "\n</w:body>\n</w:document>"
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

    # v1_explicit_single: <w:cols w:num="1"/> — parser short-circuits
    # because `if num > 1` gates ColumnLayout emission, so Page.columns
    # must stay None even though the XML element is present.
    body = _paragraph("Single column with explicit num=1.") + _sect_pr(
        '<w:cols w:num="1" w:space="0"/>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_explicit_single.docx"), body)

    # v1_three_equal: 3 equal columns with 720tw (36pt) space. No
    # equalWidth attribute → defaults to true.
    body = _paragraph("Three equal columns body.") + _sect_pr(
        '<w:cols w:num="3" w:space="720"/>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_three_equal.docx"), body)

    # v1_equal_width_false: explicit equalWidth="0" → equal_width=false.
    # Combined with 2-column num. No <w:col> children declared yet.
    body = _paragraph("Unequal-width 2-column body.") + _sect_pr(
        '<w:cols w:num="2" w:space="360" w:equalWidth="0"/>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_equal_width_false.docx"), body)

    # v1_per_column_defs: Start-element <w:cols> with child <w:col>
    # entries declaring individual widths and per-column space. Tests
    # the col_defs accumulation path on line 5546.
    cols_xml = (
        '<w:cols w:num="2" w:space="360" w:equalWidth="0">'
        '<w:col w:w="4000" w:space="240"/>'
        '<w:col w:w="5000"/>'
        '</w:cols>'
    )
    body = _paragraph("Per-column widths body.") + _sect_pr(cols_xml)
    write_docx(os.path.join(OUT_DIR, "v1_per_column_defs.docx"), body)

    # v1_no_cols: no <w:cols> at all — Page.columns must stay None
    # (the parser only emits ColumnLayout when it sees `cols`).
    body = _paragraph("No cols element body.") + _sect_pr('')
    write_docx(os.path.join(OUT_DIR, "v1_no_cols.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
