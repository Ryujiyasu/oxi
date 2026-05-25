"""Author minimal w:tabs / w:tab repro fixtures.

Paragraph.style.tab_stops end-to-end coverage (S291): parser at
parser/ooxml.rs:2422 reads `<w:tabs>` block and maps `<w:tab>` elements
(pos in twips, val for alignment, leader char) to `TabStop` IR records.
Unit tests cover XML parsing; these fixtures verify the full
parse_docx → Document walk → ParagraphStyle.tab_stops roundtrip.

Outputs to ``tools/fixtures/tabstops_samples/`` (committed, S272 no-COM
direct-write).

Fixtures (4):
  v1_simple.docx       — single left tab at 200pt
  v1_multi.docx        — left + center + right + decimal tabs
  v1_with_leader.docx  — tab with leader="dot" (TOC-style)
  v1_no_tabs.docx      — paragraph without `<w:tabs>` (tab_stops empty)
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "tabstops_samples")

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


def _tab(pos_pt: float, val: str = "left", leader: str | None = None) -> str:
    """Build a <w:tab w:val=... w:pos=... w:leader=.../> element.
    pos in points; converted to twips (twentieths of a point) internally."""
    pos_twips = int(round(pos_pt * 20))
    leader_attr = f' w:leader="{leader}"' if leader else ''
    return f'<w:tab w:val="{val}" w:pos="{pos_twips}"{leader_attr}/>'


def _paragraph_with_tabs(text: str, tabs: list[str]) -> str:
    tabs_xml = "<w:tabs>" + "".join(tabs) + "</w:tabs>" if tabs else ""
    return (
        '<w:p>'
        f'<w:pPr>{tabs_xml}</w:pPr>'
        f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
        '</w:p>'
    )


def _plain_paragraph(text: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


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

    # v1_simple: single left tab at 200pt
    body = _paragraph_with_tabs(
        "Single tab paragraph.",
        [_tab(200.0)],
    )
    write_docx(os.path.join(OUT_DIR, "v1_simple.docx"), body)

    # v1_multi: left + center + right + decimal at 100/200/300/400pt
    body = _paragraph_with_tabs(
        "Multi-tab paragraph.",
        [
            _tab(100.0, "left"),
            _tab(200.0, "center"),
            _tab(300.0, "right"),
            _tab(400.0, "decimal"),
        ],
    )
    write_docx(os.path.join(OUT_DIR, "v1_multi.docx"), body)

    # v1_with_leader: right-aligned tab with leader="dot" (typical TOC layout)
    body = _paragraph_with_tabs(
        "Chapter 1\tPage 5",
        [_tab(400.0, "right", "dot")],
    )
    write_docx(os.path.join(OUT_DIR, "v1_with_leader.docx"), body)

    # v1_no_tabs: plain paragraph, no `<w:tabs>` block → tab_stops empty
    body = _plain_paragraph("Plain paragraph, no tab stops.")
    write_docx(os.path.join(OUT_DIR, "v1_no_tabs.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
