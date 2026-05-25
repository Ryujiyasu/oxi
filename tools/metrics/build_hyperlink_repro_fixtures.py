"""Author minimal w:hyperlink repro fixtures for end-to-end parser tests.

Run.url end-to-end coverage (S272): parser/ooxml.rs:1080 wires w:hyperlink
into Run.url (either external URL via r:id → DOC_RELS, or internal anchor
via w:anchor prefixed with '#'). Unit tests cover XML parsing; these
fixtures verify the parse_docx → Document walk → Run.url roundtrip.

Outputs to ``tools/fixtures/hyperlink_samples/`` directly (committed, not via
the gitignored pipeline_data/docx/ staging — these are tiny and stable).

Fixtures (4):
  v1_external.docx       — single hyperlink to https://www.example.com/
  v1_anchor.docx         — single hyperlink to internal bookmark "section1"
  v1_mixed.docx          — plain + external + plain + anchor in one para
  v1_multirun.docx       — single hyperlink wrapping two runs (bold + plain)
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "hyperlink_samples")

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

# Fixture-specific document.xml.rels — each fixture builds its own with the
# precise set of hyperlink relationships it needs (rId3+ for w:hyperlink r:id).
def make_doc_rels(hyperlink_targets: dict[str, str]) -> str:
    """hyperlink_targets: {rId: external_url}. rIds must be >= 3 (rId1=styles, rId2=settings)."""
    extra = "\n".join(
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{escape(url)}" TargetMode="External"/>'
        for rid, url in hyperlink_targets.items()
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
{extra}
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
<w:rPrDefault>
<w:rPr>
<w:rFonts w:ascii="Calibri" w:eastAsia="ＭＳ 明朝" w:hAnsi="Calibri" w:cs="Times New Roman"/>
<w:sz w:val="22"/>
<w:szCs w:val="22"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr>
</w:rPrDefault>
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


def _run(text: str, *, bold: bool = False) -> str:
    bold_xml = '<w:rPr><w:b/></w:rPr>' if bold else ''
    return f'<w:r>{bold_xml}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def _hyperlink_external(r_id: str, *runs: str) -> str:
    """Wrap runs in w:hyperlink with r:id pointing to an external URL relationship."""
    return f'<w:hyperlink r:id="{r_id}" w:history="1">{"".join(runs)}</w:hyperlink>'


def _hyperlink_anchor(anchor: str, *runs: str) -> str:
    """Wrap runs in w:hyperlink with w:anchor (internal bookmark)."""
    return f'<w:hyperlink w:anchor="{escape(anchor)}" w:history="1">{"".join(runs)}</w:hyperlink>'


def _paragraph(*runs_or_links: str) -> str:
    return f'<w:p>{"".join(runs_or_links)}</w:p>'


def write_docx(path: str, body_xml: str, hyperlink_targets: dict[str, str]) -> None:
    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", make_doc_rels(hyperlink_targets))
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_external: single hyperlink to https://www.example.com/
    body = _paragraph(
        _run("Click here: "),
        _hyperlink_external("rId3", _run("Example")),
        _run("."),
    )
    write_docx(os.path.join(OUT_DIR, "v1_external.docx"), body,
               {"rId3": "https://www.example.com/"})

    # v1_anchor: single hyperlink to internal bookmark "section1"
    body = _paragraph(
        _run("Jump to "),
        _hyperlink_anchor("section1", _run("Section 1")),
        _run("."),
    )
    write_docx(os.path.join(OUT_DIR, "v1_anchor.docx"), body, {})

    # v1_mixed: plain + external + plain + anchor in one paragraph (3 runs become
    # 5 after parser splits hyperlink-contained runs out — 2 plain top-level +
    # 1 from external link + 1 plain + 1 from anchor link).
    body = _paragraph(
        _run("Start "),
        _hyperlink_external("rId3", _run("Ext")),
        _run(" middle "),
        _hyperlink_anchor("intro", _run("Intro")),
        _run(" end."),
    )
    write_docx(os.path.join(OUT_DIR, "v1_mixed.docx"), body,
               {"rId3": "https://docs.rs/"})

    # v1_multirun: single hyperlink wrapping TWO runs (bold + plain). Verifies
    # the parser propagates link_url to every run inside the same <w:hyperlink>.
    body = _paragraph(
        _run("Mixed-style link: "),
        _hyperlink_external("rId3",
            _run("Bold", bold=True),
            _run(" Plain"),
        ),
        _run("."),
    )
    write_docx(os.path.join(OUT_DIR, "v1_multirun.docx"), body,
               {"rId3": "https://crates.io/"})

    print("Done.")


if __name__ == "__main__":
    main()
