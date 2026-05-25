"""Author minimal w:headerReference / w:footerReference repro fixtures.

Page.header / Page.footer end-to-end coverage (S286): parser at
parser/ooxml.rs:5709 (header) / :5725 (footer) collects header/footer
references from sectPr, then reads the referenced word/headerN.xml /
word/footerN.xml parts and populates `Page.header` / `Page.footer` as
Vec<Block>. Unit tests cover XML parsing; these fixtures verify the
full parse_docx → Document walk → Page.header/footer roundtrip.

Outputs to ``tools/fixtures/header_footer_samples/`` directly (committed,
S272 no-COM direct-write variant; uses parametric XML parts like S279's
numbering).

Fixtures (4):
  v1_simple.docx       — single default header + single default footer
  v1_header_only.docx  — header only, no footer
  v1_footer_only.docx  — footer only, no header
  v1_title_page.docx   — `titlePg` + separate first-page header
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "header_footer_samples")

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


def _content_types(has_header_default: bool = False, has_header_first: bool = False,
                   has_footer_default: bool = False) -> str:
    extras = []
    if has_header_default:
        extras.append('<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>')
    if has_header_first:
        extras.append('<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>')
    if has_footer_default:
        extras.append('<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>')
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
{chr(10).join(extras)}
</Types>"""


def _doc_rels(has_header_default: bool = False, has_header_first: bool = False,
              has_footer_default: bool = False) -> str:
    extras = []
    if has_header_default:
        extras.append('<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>')
    if has_header_first:
        extras.append('<Relationship Id="rIdHdr2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>')
    if has_footer_default:
        extras.append('<Relationship Id="rIdFtr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>')
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
{chr(10).join(extras)}
</Relationships>"""


def _header_xml(text: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>
</w:hdr>"""


def _footer_xml(text: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>
</w:ftr>"""


def _sect_pr(header_refs: list = None, footer_refs: list = None, title_pg: bool = False) -> str:
    header_refs = header_refs or []  # list of (type, rel_id)
    footer_refs = footer_refs or []
    refs = ""
    for ref_type, rel_id in header_refs:
        refs += f'<w:headerReference w:type="{ref_type}" r:id="{rel_id}"/>'
    for ref_type, rel_id in footer_refs:
        refs += f'<w:footerReference w:type="{ref_type}" r:id="{rel_id}"/>'
    title_pg_xml = '<w:titlePg/>' if title_pg else ''
    return (
        '<w:sectPr>'
        + refs +
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        + title_pg_xml +
        '</w:sectPr>'
    )


def _body_para(text: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def write_docx(path: str, body_xml: str, parts: dict[str, str]) -> None:
    """parts: extra zip entries beyond the standard 5 (e.g. header1.xml, footer1.xml).
    Standard parts ([Content_Types].xml, rels, document.xml, styles.xml,
    settings.xml) are written by this function based on `parts` keys.
    """
    has_hdr1 = "word/header1.xml" in parts
    has_hdr2 = "word/header2.xml" in parts
    has_ftr1 = "word/footer1.xml" in parts
    full_doc = DOC_HEAD + body_xml + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _content_types(has_hdr1, has_hdr2, has_ftr1))
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", _doc_rels(has_hdr1, has_hdr2, has_ftr1))
        z.writestr("word/document.xml", full_doc)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
        for name, content in parts.items():
            z.writestr(name, content)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_simple: 1 body paragraph + default header + default footer
    sect = _sect_pr(
        header_refs=[("default", "rIdHdr1")],
        footer_refs=[("default", "rIdFtr1")],
    )
    body = _body_para("Body text on page 1.") + sect
    write_docx(
        os.path.join(OUT_DIR, "v1_simple.docx"),
        body,
        {
            "word/header1.xml": _header_xml("Header content - default"),
            "word/footer1.xml": _footer_xml("Footer content - default"),
        },
    )

    # v1_header_only: header but no footer
    sect = _sect_pr(header_refs=[("default", "rIdHdr1")])
    body = _body_para("Body text only, no footer.") + sect
    write_docx(
        os.path.join(OUT_DIR, "v1_header_only.docx"),
        body,
        {"word/header1.xml": _header_xml("Header only — no footer")},
    )

    # v1_footer_only: footer but no header
    sect = _sect_pr(footer_refs=[("default", "rIdFtr1")])
    body = _body_para("Body text only, no header.") + sect
    write_docx(
        os.path.join(OUT_DIR, "v1_footer_only.docx"),
        body,
        {"word/footer1.xml": _footer_xml("Footer only — no header")},
    )

    # v1_title_page: titlePg + separate first-page header
    sect = _sect_pr(
        header_refs=[("default", "rIdHdr1"), ("first", "rIdHdr2")],
        footer_refs=[("default", "rIdFtr1")],
        title_pg=True,
    )
    body = _body_para("Body text on title page.") + sect
    write_docx(
        os.path.join(OUT_DIR, "v1_title_page.docx"),
        body,
        {
            "word/header1.xml": _header_xml("Default header (non-title pages)"),
            "word/header2.xml": _header_xml("First-page-only header"),
            "word/footer1.xml": _footer_xml("Default footer"),
        },
    )

    print("Done.")


if __name__ == "__main__":
    main()
