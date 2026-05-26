"""Author minimal w:pBdr / w:top/bottom/left/right/between repro fixtures.

Paragraph.style.borders end-to-end coverage (S302): parser at
parser/ooxml.rs:1915 routes <w:pBdr> to parse_paragraph_borders
(parser/ooxml.rs:2362), which reads top/bottom/left/right/between
children and maps each to Option<BorderDef> via parse_border_attrs
(parser/ooxml.rs:2394):
  - w:val="none"|"nil" → None (suppressed)
  - w:val=<style>      → BorderDef.style verbatim ("single", "dashed", ...)
  - w:sz=<eighths-pt>  → BorderDef.width = sz/8 in pt
  - w:color="auto"     → BorderDef.color = "000000"
  - w:color=<hex>      → BorderDef.color = hex verbatim
  - w:start / w:end    → mapped to left / right (newer OOXML alias)

Unit tests cover XML parsing in isolation; these fixtures verify the
full parse_docx → Paragraph.style.borders roundtrip including
inheritance from doc defaults and the start/end alias mapping.

Outputs to ``tools/fixtures/paragraph_borders_samples/`` (committed,
S272 no-COM direct-write pattern).

Fixtures (5):
  v1_all_sides.docx       — top + bottom + left + right (4 borders, single)
  v1_between.docx         — <w:between> border surfaces in IR
  v1_start_end_aliases.docx — uses <w:start>/<w:end> → maps to left/right
  v1_color_and_width.docx — sz=12→1.5pt, color hex preserved, color=auto→000000
  v1_none_suppresses.docx — val="none" and val="nil" yield None
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "paragraph_borders_samples")

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


def _border(local: str, val: str = "single", sz: int = 4,
            space: int = 0, color: str | None = "000000") -> str:
    """Build a <w:top/bottom/left/right/start/end/between> element.
       sz in eighths of a pt; sz=4 → 0.5pt."""
    color_attr = f' w:color="{color}"' if color is not None else ''
    return (
        f'<w:{local} w:val="{val}" w:sz="{sz}" w:space="{space}"{color_attr}/>'
    )


def _paragraph_with_borders(text: str, borders_xml: str) -> str:
    pbdr = f'<w:pBdr>{borders_xml}</w:pBdr>' if borders_xml else ''
    return (
        '<w:p>'
        f'<w:pPr>{pbdr}</w:pPr>'
        f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
        '</w:p>'
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

    # v1_all_sides: 4 sides, single style, sz=4 (0.5pt), color=000000
    body = _paragraph_with_borders(
        "Para with top+bottom+left+right borders.",
        _border("top") + _border("bottom") + _border("left") + _border("right"),
    )
    write_docx(os.path.join(OUT_DIR, "v1_all_sides.docx"), body)

    # v1_between: between border for paragraphs sharing same pBdr
    body = _paragraph_with_borders(
        "Para with <w:between> border.",
        _border("between"),
    )
    write_docx(os.path.join(OUT_DIR, "v1_between.docx"), body)

    # v1_start_end_aliases: <w:start>/<w:end> map to left/right
    body = _paragraph_with_borders(
        "Para using <w:start>/<w:end> aliases.",
        _border("start") + _border("end"),
    )
    write_docx(os.path.join(OUT_DIR, "v1_start_end_aliases.docx"), body)

    # v1_color_and_width: sz=12 (1.5pt), color="FF0000", color="auto"
    body = _paragraph_with_borders(
        "Para with 1.5pt red top and auto-color bottom.",
        _border("top", sz=12, color="FF0000")
        + _border("bottom", sz=4, color="auto"),
    )
    write_docx(os.path.join(OUT_DIR, "v1_color_and_width.docx"), body)

    # v1_none_suppresses: val="none" and val="nil" → None
    body = _paragraph_with_borders(
        'Para with val="none"/"nil" suppressors.',
        _border("top", val="none") + _border("bottom", val="nil"),
    )
    write_docx(os.path.join(OUT_DIR, "v1_none_suppresses.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
