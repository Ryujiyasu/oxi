"""Author minimal w:footnoteReference / w:endnoteReference repro fixtures.

Run.footnote_ref / Run.endnote_ref end-to-end coverage (S273): parser at
parser/ooxml.rs:2602 (footnote) / 2619 (endnote) extracts the w:id and
stores it on Run; later renumber_note_refs (parser/ooxml.rs:5907) rewrites
run.text to the per-section sequence number ("1","2","3",...) regardless of
the raw XML id (which may start at 2 or have gaps because id=1 is reserved
for the separator).

Outputs to ``tools/fixtures/footnote_samples/`` directly (committed; no
COM-measurement step needed for parser-only assertions, matching the S272
hyperlink direct-write variant).

Fixtures (4):
  v1_footnote.docx       — single footnote (id=1) → text "1"
  v1_endnote.docx        — single endnote   (id=1) → text "1"
  v1_mixed.docx          — one footnote + one endnote in same paragraph
  v1_renumber.docx       — two footnotes with raw ids 2 and 5 (renumber → 1, 2)
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "footnote_samples")

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
"""

SECT_PR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '</w:sectPr>'
)

FOOTNOTES_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
"""

FOOTNOTES_TAIL = "\n</w:footnotes>"

ENDNOTES_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
"""

ENDNOTES_TAIL = "\n</w:endnotes>"


def _content_types(has_footnotes: bool, has_endnotes: bool) -> str:
    extras = []
    if has_footnotes:
        extras.append('<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>')
    if has_endnotes:
        extras.append('<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>')
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
{chr(10).join(extras)}
</Types>"""


def _doc_rels(has_footnotes: bool, has_endnotes: bool) -> str:
    extras = []
    if has_footnotes:
        extras.append('<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>')
    if has_endnotes:
        extras.append('<Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>')
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
{chr(10).join(extras)}
</Relationships>"""


def _run_text(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def _footnote_ref_run(fn_id: int) -> str:
    return (
        '<w:r>'
        '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
        f'<w:footnoteReference w:id="{fn_id}"/>'
        '</w:r>'
    )


def _endnote_ref_run(en_id: int) -> str:
    return (
        '<w:r>'
        '<w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>'
        f'<w:endnoteReference w:id="{en_id}"/>'
        '</w:r>'
    )


def _footnote_def(fn_id: int, text: str) -> str:
    return (
        f'<w:footnote w:id="{fn_id}">'
        '<w:p>'
        '<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
        f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
        '</w:p>'
        '</w:footnote>'
    )


def _endnote_def(en_id: int, text: str) -> str:
    return (
        f'<w:endnote w:id="{en_id}">'
        '<w:p>'
        '<w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>'
        '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>'
        f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
        '</w:p>'
        '</w:endnote>'
    )


def _paragraph(*runs: str) -> str:
    return f'<w:p>{"".join(runs)}</w:p>'


def write_docx(
    path: str,
    body_xml: str,
    *,
    footnote_defs: list[str] = None,
    endnote_defs: list[str] = None,
) -> None:
    footnote_defs = footnote_defs or []
    endnote_defs = endnote_defs or []
    has_fn = bool(footnote_defs)
    has_en = bool(endnote_defs)

    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _content_types(has_fn, has_en))
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", _doc_rels(has_fn, has_en))
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
        if has_fn:
            z.writestr("word/footnotes.xml", FOOTNOTES_HEAD + "\n".join(footnote_defs) + FOOTNOTES_TAIL)
        if has_en:
            z.writestr("word/endnotes.xml", ENDNOTES_HEAD + "\n".join(endnote_defs) + ENDNOTES_TAIL)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_footnote: single footnote with raw id=1
    body = _paragraph(
        _run_text("Body text with footnote"),
        _footnote_ref_run(1),
        _run_text(" and more body text."),
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_footnote.docx"),
        body,
        footnote_defs=[_footnote_def(1, " First footnote content.")],
    )

    # v1_endnote: single endnote with raw id=1
    body = _paragraph(
        _run_text("Body text with endnote"),
        _endnote_ref_run(1),
        _run_text(" and more body text."),
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_endnote.docx"),
        body,
        endnote_defs=[_endnote_def(1, " First endnote content.")],
    )

    # v1_mixed: footnote + endnote in same paragraph
    body = _paragraph(
        _run_text("See footnote"),
        _footnote_ref_run(1),
        _run_text(" and endnote"),
        _endnote_ref_run(1),
        _run_text("."),
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_mixed.docx"),
        body,
        footnote_defs=[_footnote_def(1, " The footnote.")],
        endnote_defs=[_endnote_def(1, " The endnote.")],
    )

    # v1_renumber: footnotes with raw ids 2 and 5 (Word allows gaps; parser
    # renumbers them to sequence "1", "2" in run.text but preserves raw id
    # in run.footnote_ref).
    body = _paragraph(
        _run_text("First marker"),
        _footnote_ref_run(2),
        _run_text(" then second marker"),
        _footnote_ref_run(5),
        _run_text("."),
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_renumber.docx"),
        body,
        footnote_defs=[
            _footnote_def(2, " Note A (raw id=2)."),
            _footnote_def(5, " Note B (raw id=5)."),
        ],
    )

    print("Done.")


if __name__ == "__main__":
    main()
