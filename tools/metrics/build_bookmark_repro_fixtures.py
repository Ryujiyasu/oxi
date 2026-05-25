"""Author minimal w:bookmarkStart / w:bookmarkEnd repro fixtures.

Run.bookmark_name end-to-end coverage (S274): parser at parser/ooxml.rs:1341
turns each `<w:bookmarkStart w:name="X"/>` into an empty anchor Run with
bookmark_name=Some("X"). `_GoBack` (which Word auto-inserts to remember the
cursor position) is filtered at parser/ooxml.rs:1347 and produces no run.
`<w:bookmarkEnd>` is a no-op (anchor is placed at start).

Outputs to ``tools/fixtures/bookmark_samples/`` directly (committed, no-COM
direct-write variant per S272/S273).

Fixtures (4):
  v1_basic.docx           — single bookmark "section1" wrapping plain text
  v1_goback_skipped.docx  — explicit _GoBack bookmark must produce NO run
  v1_multiple.docx        — three bookmarks (intro / body / end) in one para
  v1_around_text.docx     — bookmark wraps text; anchor is at START, text follows
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "bookmark_samples")

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


def _bookmark_start(bk_id: int, name: str) -> str:
    return f'<w:bookmarkStart w:id="{bk_id}" w:name="{escape(name)}"/>'


def _bookmark_end(bk_id: int) -> str:
    return f'<w:bookmarkEnd w:id="{bk_id}"/>'


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

    # v1_basic: bookmark "section1" placed before text content
    body = _paragraph(
        _bookmark_start(1, "section1"),
        _run("Hello, world."),
        _bookmark_end(1),
    )
    write_docx(os.path.join(OUT_DIR, "v1_basic.docx"), body)

    # v1_goback_skipped: _GoBack is Word's auto-inserted cursor bookmark.
    # Parser at parser/ooxml.rs:1347 explicitly filters it. Verifies the
    # filter is active end-to-end — only the "real" bookmark survives.
    body = _paragraph(
        _bookmark_start(1, "_GoBack"),
        _run("This text has a _GoBack marker. "),
        _bookmark_end(1),
        _bookmark_start(2, "real_anchor"),
        _run("This part has a real anchor."),
        _bookmark_end(2),
    )
    write_docx(os.path.join(OUT_DIR, "v1_goback_skipped.docx"), body)

    # v1_multiple: three distinct bookmarks intro / body / end
    body = _paragraph(
        _bookmark_start(1, "intro"),
        _run("Intro. "),
        _bookmark_end(1),
        _bookmark_start(2, "body"),
        _run("Body. "),
        _bookmark_end(2),
        _bookmark_start(3, "end"),
        _run("End."),
        _bookmark_end(3),
    )
    write_docx(os.path.join(OUT_DIR, "v1_multiple.docx"), body)

    # v1_around_text: bookmark wrapping text — bookmarkStart precedes the run,
    # bookmarkEnd follows it. Parser collapses start to an empty anchor run AT
    # THE START position (no end-anchor); text run stays untouched after it.
    body = _paragraph(
        _run("Before. "),
        _bookmark_start(1, "wrap"),
        _run("Inside wrapped span."),
        _bookmark_end(1),
        _run(" After."),
    )
    write_docx(os.path.join(OUT_DIR, "v1_around_text.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
