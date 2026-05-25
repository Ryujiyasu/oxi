"""Author minimal w:numPr repro fixtures backed by word/numbering.xml.

Paragraph.style.{num_id, num_ilvl, list_marker, list_suff} end-to-end
coverage (S279): parser at parser/ooxml.rs:1599 resolves num_id+ilvl
against the doc's numbering definitions and populates the resolved
list_marker / list_suff / list_indent fields. Unit tests cover the
resolve_marker_full path; these fixtures verify the parse_docx →
Document walk → ParagraphStyle.list_* roundtrip.

Outputs to ``tools/fixtures/numbering_samples/`` directly (committed;
S272 no-COM direct-write variant).

Fixtures (4):
  v1_decimal.docx          — 3 paragraphs at ilvl=0, decimal numbering "1.","2.","3."
  v1_bullet.docx           — 3 bullet paragraphs at ilvl=0 ("•")
  v1_two_levels.docx       — mixed ilvl=0 + ilvl=1 (nested list)
  v1_two_numids.docx       — two independent numId sequences (each restarts at 1)
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "numbering_samples")

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
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


def _numbering_xml(*abstract_defs: tuple[int, list[tuple[int, str, str]]],
                   nums: list[tuple[int, int]]) -> str:
    """Build word/numbering.xml.

    abstract_defs: each (abstract_num_id, [(ilvl, num_fmt, lvl_text), ...])
    nums:          [(num_id, abstract_num_id), ...]
    """
    abstracts_xml = ""
    for abs_id, levels in abstract_defs:
        levels_xml = ""
        for ilvl, num_fmt, lvl_text in levels:
            levels_xml += (
                f'<w:lvl w:ilvl="{ilvl}">'
                '<w:start w:val="1"/>'
                f'<w:numFmt w:val="{num_fmt}"/>'
                f'<w:lvlText w:val="{escape(lvl_text)}"/>'
                '<w:lvlJc w:val="left"/>'
                '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
                '</w:lvl>'
            )
        abstracts_xml += f'<w:abstractNum w:abstractNumId="{abs_id}">{levels_xml}</w:abstractNum>'

    nums_xml = ""
    for num_id, abs_id in nums:
        nums_xml += (
            f'<w:num w:numId="{num_id}">'
            f'<w:abstractNumId w:val="{abs_id}"/>'
            '</w:num>'
        )

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
{abstracts_xml}{nums_xml}
</w:numbering>"""


def _numbered_p(text: str, num_id: int, ilvl: int = 0) -> str:
    return (
        '<w:p>'
        '<w:pPr>'
        f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr>'
        '</w:pPr>'
        f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
        '</w:p>'
    )


def write_docx(path: str, body_xml: str, numbering_xml: str) -> None:
    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
        z.writestr("word/numbering.xml", numbering_xml)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_decimal: 3 numbered paragraphs at ilvl=0 with decimal numbering
    # The lvlText "%1." is Word's placeholder syntax: %1 = level 1 counter.
    numbering = _numbering_xml(
        (0, [(0, "decimal", "%1.")]),
        nums=[(1, 0)],
    )
    body = (
        _numbered_p("Alpha", num_id=1)
        + _numbered_p("Bravo", num_id=1)
        + _numbered_p("Charlie", num_id=1)
    )
    write_docx(os.path.join(OUT_DIR, "v1_decimal.docx"), body, numbering)

    # v1_bullet: 3 bullet paragraphs
    numbering = _numbering_xml(
        (0, [(0, "bullet", "•")]),  # • bullet
        nums=[(1, 0)],
    )
    body = (
        _numbered_p("Apple", num_id=1)
        + _numbered_p("Banana", num_id=1)
        + _numbered_p("Cherry", num_id=1)
    )
    write_docx(os.path.join(OUT_DIR, "v1_bullet.docx"), body, numbering)

    # v1_two_levels: nested numbering — ilvl=0 (decimal) + ilvl=1 (lower-alpha)
    numbering = _numbering_xml(
        (0, [
            (0, "decimal", "%1."),
            (1, "lowerLetter", "%2)"),
        ]),
        nums=[(1, 0)],
    )
    body = (
        _numbered_p("Top 1", num_id=1, ilvl=0)
        + _numbered_p("Nested 1a", num_id=1, ilvl=1)
        + _numbered_p("Nested 1b", num_id=1, ilvl=1)
        + _numbered_p("Top 2", num_id=1, ilvl=0)
    )
    write_docx(os.path.join(OUT_DIR, "v1_two_levels.docx"), body, numbering)

    # v1_two_numids: independent numId sequences should each start at 1
    numbering = _numbering_xml(
        (0, [(0, "decimal", "%1.")]),
        (1, [(0, "decimal", "%1.")]),
        nums=[(1, 0), (2, 1)],
    )
    body = (
        _numbered_p("List A first",  num_id=1)
        + _numbered_p("List A second", num_id=1)
        + _numbered_p("List B first",  num_id=2)
        + _numbered_p("List B second", num_id=2)
        + _numbered_p("List A third",  num_id=1)
    )
    write_docx(os.path.join(OUT_DIR, "v1_two_numids.docx"), body, numbering)

    print("Done.")


if __name__ == "__main__":
    main()
