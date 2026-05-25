"""Author minimal sectPr (page size / margins / columns / orientation) repro fixtures.

Page.size / Page.margin / Page.columns end-to-end coverage (S290):
parser at parser/ooxml.rs:5571 (pgSz) / :5596 (pgMar) / :5766 (cols)
reads section properties and populates the per-Page IR fields. Unit
tests cover XML-level parsing; these fixtures verify the full
parse_docx → Document walk → Page.{size, margin, columns} roundtrip
across the common section property combinations.

Outputs to ``tools/fixtures/section_samples/`` directly (committed,
S272 no-COM direct-write variant).

Fixtures (4):
  v1_a4_portrait.docx     — standard A4, 1-inch margins, 1 column
  v1_a4_landscape.docx    — A4 with orient="landscape" (width/height swap)
  v1_custom_margins.docx  — non-default margins
                            (top=100pt, bottom=50pt, left=80pt, right=40pt)
  v1_two_columns.docx     — 2-column layout with spacing
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "section_samples")

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


def _sect_pr(pg_w: int, pg_h: int, orient: str | None,
             margins: dict[str, int], cols: dict | None = None) -> str:
    """Build a sectPr. All measurements in twentieths of a point (twips * 0.05).

    `margins` keys: top, right, bottom, left, header, footer, gutter
    `cols` if present: {'num': int, 'space': int}
    """
    orient_attr = f' w:orient="{orient}"' if orient else ''
    cols_xml = ''
    if cols:
        attrs = ' '.join(f'w:{k}="{v}"' for k, v in cols.items())
        cols_xml = f'<w:cols {attrs}/>'
    m = margins
    return (
        '<w:sectPr>'
        f'<w:pgSz w:w="{pg_w}" w:h="{pg_h}"{orient_attr}/>'
        f'<w:pgMar w:top="{m["top"]}" w:right="{m["right"]}" '
        f'w:bottom="{m["bottom"]}" w:left="{m["left"]}" '
        f'w:header="{m.get("header", 851)}" w:footer="{m.get("footer", 992)}" '
        f'w:gutter="{m.get("gutter", 0)}"/>'
        + cols_xml +
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

    # v1_a4_portrait: A4 (11906×16838 twips = 595.3×841.9pt),
    # standard Word default margins (1440 twips top/bottom/left/right = 72pt = 1inch).
    body = _paragraph("Standard A4 portrait body.") + _sect_pr(
        pg_w=11906, pg_h=16838, orient=None,
        margins=dict(top=1440, right=1440, bottom=1440, left=1440)
    )
    write_docx(os.path.join(OUT_DIR, "v1_a4_portrait.docx"), body)

    # v1_a4_landscape: same A4 dims but orient="landscape" → parser swaps
    # width and height. Word's XML keeps the values as-is (w smaller than h)
    # and uses orient attribute to indicate the actual display orientation.
    body = _paragraph("A4 landscape body.") + _sect_pr(
        pg_w=11906, pg_h=16838, orient="landscape",
        margins=dict(top=1440, right=1440, bottom=1440, left=1440)
    )
    write_docx(os.path.join(OUT_DIR, "v1_a4_landscape.docx"), body)

    # v1_custom_margins: A4 with non-default margins.
    # 100pt top = 2000 twips, 50pt bottom = 1000, 80pt left = 1600, 40pt right = 800.
    body = _paragraph("Custom margins body.") + _sect_pr(
        pg_w=11906, pg_h=16838, orient=None,
        margins=dict(top=2000, right=800, bottom=1000, left=1600)
    )
    write_docx(os.path.join(OUT_DIR, "v1_custom_margins.docx"), body)

    # v1_two_columns: 2-column layout with 360twips (18pt) spacing between.
    body = _paragraph("Two column body.") + _sect_pr(
        pg_w=11906, pg_h=16838, orient=None,
        margins=dict(top=1440, right=1440, bottom=1440, left=1440),
        cols=dict(num=2, space=360)
    )
    write_docx(os.path.join(OUT_DIR, "v1_two_columns.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
