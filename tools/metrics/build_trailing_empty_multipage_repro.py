"""Multi-page extension of build_trailing_empty_cell_repro (S304).

Single-page repro (v1/v2) shows Oxi reserves 11.6pt for the trailing
empty — matches Word. But e3c545 cell content OVERFLOWS to a second
page, with the trailing empty on the second page. Test whether Oxi
loses the trailing empty's height when the cell spans pages.

Each variant: 1-cell table with N text paragraphs (enough to overflow)
+ K trailing empty paragraphs, followed by a body marker.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "trailing_empty_multipage_repro")

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
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="ＭＳ 明朝" w:hAnsi="Calibri"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
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


def _cell_para_text(text: str) -> str:
    return (
        '<w:p>'
        '<w:pPr>'
        '<w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック"/>'
        '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        '</w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック"/>'
        '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(text)}</w:t>'
        '</w:r>'
        '</w:p>'
    )


def _cell_para_empty() -> str:
    return (
        '<w:p>'
        '<w:pPr>'
        '<w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック"/>'
        '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        '</w:pPr>'
        '</w:p>'
    )


def _body_para(text: str) -> str:
    return (
        f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
    )


def _wrap_table(cell_paras: str) -> str:
    return (
        '<w:tbl>'
        '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
        '<w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>'
        f'{cell_paras}'
        '</w:tc>'
        '</w:tr>'
        '</w:tbl>'
    )


def build_doc(n_text_paras: int, n_trailing_empty: int) -> str:
    """Build a doc with N text + K trailing empty paragraphs in a single
    1-cell 1-row table. n_text_paras is sized so the cell overflows
    to a second page (page height ~11.7in, ~58 lines at 9pt+1.4 spacing)."""
    cell_paras = ""
    for i in range(n_text_paras):
        cell_paras += _cell_para_text(f"line {i+1:03d} content")
    for _ in range(n_trailing_empty):
        cell_paras += _cell_para_empty()
    body = (
        _body_para("BEFORE TABLE marker.")
        + _wrap_table(cell_paras)
        + _body_para("AFTER TABLE marker.")
    )
    return DOC_HEAD + body + SECT_PR + "\n</w:body>\n</w:document>"


def write_docx(path: str, doc_xml: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")
    # Need enough text to span 2 pages. At 9pt with 12pt line height,
    # a page can fit ~58 lines. Use 70 to ensure overflow.
    NTEXT = 70
    write_docx(os.path.join(OUT_DIR, "v1_multipage_no_trailing.docx"),
               build_doc(NTEXT, 0))
    write_docx(os.path.join(OUT_DIR, "v2_multipage_one_trailing.docx"),
               build_doc(NTEXT, 1))
    write_docx(os.path.join(OUT_DIR, "v3_multipage_two_trailing.docx"),
               build_doc(NTEXT, 2))
    print("Done.")


if __name__ == "__main__":
    main()
