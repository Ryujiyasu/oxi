"""Author minimal w:fldChar / w:instrText repro fixtures.

Run.field_type end-to-end coverage (S278): parser at parser/ooxml.rs:2656
inspects the `<w:instrText>` content of a complete field (delimited by
<w:fldChar fldCharType="begin"/.../>...separate/.../>...end/>) and sets
Run.field_type when it recognizes PAGE or NUMPAGES. Other field codes
(DATE, TOC, HYPERLINK, REF, AUTHOR/TITLE/SUBJECT) are handled with
text-only rewrite and no field_type variant.

Outputs to ``tools/fixtures/field_samples/`` directly (committed; no COM
needed for parser-only assertions, S272 direct-write variant).

Fixtures (4):
  v1_page.docx              — `<w:instrText>PAGE \\* MERGEFORMAT</w:instrText>` → field_type=Some(Page), text="#"
  v1_numpages.docx          — NUMPAGES → field_type=Some(NumPages), text="#"
  v1_page_of_numpages.docx  — common "Page # of #" header pattern in 1 paragraph
  v1_date.docx              — DATE → no field_type variant, text contains "DATE"
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "field_samples")

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


def _run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def _field(instr_text: str, result_text: str = "1") -> str:
    """Complete field group: begin → instrText → separate → result → end.

    Parser collapses begin/separate/end markers into U+FFFE / U+FFFF in
    the run text stream so parse_paragraph can track field-result depth,
    and reads the instrText to set Run.field_type and rewrite the result
    run's text.
    """
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve">{escape(instr_text)}</w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:t xml:space="preserve">{escape(result_text)}</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def _paragraph(*children: str) -> str:
    return f'<w:p>{"".join(children)}</w:p>'


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

    # v1_page: a single PAGE field with the typical "* MERGEFORMAT" switch.
    body = _paragraph(
        _run("Page "),
        _field("PAGE   \\* MERGEFORMAT", result_text="1"),
        _run(" of total."),
    )
    write_docx(os.path.join(OUT_DIR, "v1_page.docx"), body)

    # v1_numpages: NUMPAGES alone.
    body = _paragraph(
        _run("Total: "),
        _field("NUMPAGES   \\* MERGEFORMAT", result_text="42"),
        _run(" pages."),
    )
    write_docx(os.path.join(OUT_DIR, "v1_numpages.docx"), body)

    # v1_page_of_numpages: common header "Page # of #" pattern — PAGE and
    # NUMPAGES in the same paragraph. Verifies both field_type variants
    # coexist and don't clobber each other.
    body = _paragraph(
        _run("Page "),
        _field("PAGE", result_text="3"),
        _run(" of "),
        _field("NUMPAGES", result_text="10"),
        _run("."),
    )
    write_docx(os.path.join(OUT_DIR, "v1_page_of_numpages.docx"), body)

    # v1_date: DATE field — parser rewrites text to the field-code text
    # but does NOT set Run.field_type (only Page/NumPages do).
    body = _paragraph(
        _run("Today: "),
        _field("DATE \\@ \"yyyy/M/d\"", result_text="2026/5/25"),
        _run("."),
    )
    write_docx(os.path.join(OUT_DIR, "v1_date.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
