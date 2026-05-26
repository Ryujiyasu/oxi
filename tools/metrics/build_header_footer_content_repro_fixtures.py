"""Author minimal `word/headerN.xml` / `word/footerN.xml` repro fixtures
for S319 — header/footer INNER CONTENT coverage.

`header_footer_integration.rs` (S286) covers the OUTER routing:
title_pg dispatch, default/first ref type, presence/absence of
Page.header/Page.footer Vec<Block>. But the INNER content parsing
in `parse_header_footer_xml` (parser/ooxml.rs:5835) was not pinned:

  - Multiple <w:p> children inside <w:hdr> → multiple Block::Paragraph
    entries in Page.header in source order.
  - <w:tbl> inside <w:hdr> → Block::Table at the same level as
    Block::Paragraph (line 5854-5857).
  - <w:sdt> wraps <w:sdtPr> (skipped) + <w:sdtContent> (harvested).
    Paragraphs and tables INSIDE <w:sdtContent> appear directly in
    Page.header — the sdt wrapper itself is NOT a Block (line
    5858-5891). NON-OBVIOUS: a regression that stored sdt as its
    own block, or that dropped sdt content entirely, would silently
    affect Word docs that use content controls in headers (common
    in templates).
  - Paragraph properties (pPr/jc) and run properties (rPr/b) inside
    header paragraphs MUST propagate end-to-end. The header parser
    delegates to parse_paragraph (line 5851) so the same property
    handling applies — pinning this catches an accidental override
    where header parsing forgot to plumb ctx/styles.

Outputs to ``tools/fixtures/header_footer_content_samples/``.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "header_footer_content_samples")


def _content_types(has_hdr: bool = False, has_ftr: bool = False) -> str:
    extras = []
    if has_hdr:
        extras.append('<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>')
    if has_ftr:
        extras.append('<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '<Default Extension="xml" ContentType="application/xml"/>\n'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
        + '\n'.join(extras) +
        '\n</Types>'
    )


def _doc_rels(has_hdr: bool = False, has_ftr: bool = False) -> str:
    extras = []
    if has_hdr:
        extras.append('<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>')
    if has_ftr:
        extras.append('<Relationship Id="rIdFtr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n'
        + '\n'.join(extras) +
        '\n</Relationships>'
    )


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
<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>
"""


def _sect_pr_with_hdr_ftr(hdr_rel: str = "", ftr_rel: str = "") -> str:
    refs = ""
    if hdr_rel:
        refs += f'<w:headerReference w:type="default" r:id="{hdr_rel}"/>'
    if ftr_rel:
        refs += f'<w:footerReference w:type="default" r:id="{ftr_rel}"/>'
    return (
        '<w:sectPr>' + refs +
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '</w:sectPr>'
    )


def _hdr_wrap(inner: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        + inner +
        '\n</w:hdr>'
    )


def _ftr_wrap(inner: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        + inner +
        '\n</w:ftr>'
    )


def _para(text: str, ppr_xml: str = "", rpr_xml: str = "") -> str:
    ppr = f'<w:pPr>{ppr_xml}</w:pPr>' if ppr_xml else ''
    rpr = f'<w:rPr>{rpr_xml}</w:rPr>' if rpr_xml else ''
    return (
        f'<w:p>{ppr}<w:r>{rpr}'
        f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
    )


def write_docx(path: str, sect_pr_xml: str, parts: dict) -> None:
    has_hdr = "word/header1.xml" in parts
    has_ftr = "word/footer1.xml" in parts
    full_doc = DOC_HEAD + sect_pr_xml + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _content_types(has_hdr, has_ftr))
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels",
                   _doc_rels(has_hdr, has_ftr))
        z.writestr("word/document.xml", full_doc)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
        for name, content in parts.items():
            z.writestr(name, content)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_header_two_paragraphs_order: header with 2 paragraphs.
    # Pin Page.header[0] = "first" and Page.header[1] = "second"
    # (source order preserved through parse_header_footer_xml).
    hdr = _hdr_wrap(_para("first") + _para("second"))
    sect_pr = _sect_pr_with_hdr_ftr(hdr_rel="rIdHdr1")
    write_docx(
        os.path.join(OUT_DIR, "v1_header_two_paragraphs_order.docx"),
        sect_pr,
        {"word/header1.xml": hdr},
    )

    # v1_header_table: header with 1×2 table. Pin Block::Table
    # routing inside Page.header (parser/ooxml.rs:5854-5857).
    tbl_xml = (
        '<w:tbl>'
        '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>'
        '<w:tr>'
        '<w:tc><w:tcPr/>' + _para("hdr-cell-A") + '</w:tc>'
        '<w:tc><w:tcPr/>' + _para("hdr-cell-B") + '</w:tc>'
        '</w:tr>'
        '</w:tbl>'
    )
    hdr = _hdr_wrap(tbl_xml)
    sect_pr = _sect_pr_with_hdr_ftr(hdr_rel="rIdHdr1")
    write_docx(
        os.path.join(OUT_DIR, "v1_header_table.docx"),
        sect_pr,
        {"word/header1.xml": hdr},
    )

    # v1_header_sdt_paragraph: <w:sdt> wraps a <w:p>. The sdt
    # wrapper itself is NOT a Block — its inner paragraph is
    # harvested directly into Page.header (parser/ooxml.rs:
    # 5858-5891). NON-OBVIOUS: a regression that stored sdt as
    # its own block would silently affect content-control headers.
    sdt = (
        '<w:sdt>'
        '<w:sdtPr><w:id w:val="100"/></w:sdtPr>'
        '<w:sdtContent>'
        + _para("inside-sdt") +
        '</w:sdtContent>'
        '</w:sdt>'
    )
    hdr = _hdr_wrap(sdt)
    sect_pr = _sect_pr_with_hdr_ftr(hdr_rel="rIdHdr1")
    write_docx(
        os.path.join(OUT_DIR, "v1_header_sdt_paragraph.docx"),
        sect_pr,
        {"word/header1.xml": hdr},
    )

    # v1_header_sdt_table: <w:sdt> wraps a <w:tbl>. Pins the OTHER
    # sdtContent branch (line 5871-5874) — table inside sdt is
    # still routed to Block::Table at the same level as a
    # non-wrapped table.
    sdt = (
        '<w:sdt>'
        '<w:sdtPr><w:id w:val="200"/></w:sdtPr>'
        '<w:sdtContent>'
        + tbl_xml +
        '</w:sdtContent>'
        '</w:sdt>'
    )
    hdr = _hdr_wrap(sdt)
    sect_pr = _sect_pr_with_hdr_ftr(hdr_rel="rIdHdr1")
    write_docx(
        os.path.join(OUT_DIR, "v1_header_sdt_table.docx"),
        sect_pr,
        {"word/header1.xml": hdr},
    )

    # v1_footer_para_with_properties: footer paragraph with
    # pPr/jc=center + rPr/b. Pins that parse_paragraph runs with
    # full property handling inside parse_header_footer_xml — the
    # header/footer parser delegates to parse_paragraph (line 5851)
    # so pPr / rPr must propagate end-to-end.
    ftr = _ftr_wrap(_para(
        "centered-bold",
        ppr_xml='<w:jc w:val="center"/>',
        rpr_xml='<w:b/>',
    ))
    sect_pr = _sect_pr_with_hdr_ftr(ftr_rel="rIdFtr1")
    write_docx(
        os.path.join(OUT_DIR, "v1_footer_para_with_properties.docx"),
        sect_pr,
        {"word/footer1.xml": ftr},
    )

    print("Done.")


if __name__ == "__main__":
    main()
