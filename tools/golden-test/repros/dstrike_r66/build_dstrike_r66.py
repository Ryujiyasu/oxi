"""R66 minimal repro — w:dstrike (double strikethrough) coverage feature.

Generates a single .docx exercising w:dstrike in multiple contexts:
  L1 plain         — control (no decoration)
  L2 single strike — w:strike
  L3 double strike — w:dstrike  (R66 target)
  L4 dstrike+bold  — w:dstrike with w:b
  L5 dstrike CJK   — w:dstrike on MS Mincho fullwidth glyphs
  L6 dstrike color — w:dstrike with w:color w:val="FF0000"

Word renders w:dstrike as two parallel horizontal lines through the
glyph midline. Oxi must do the same. Coverage ship per CLAUDE.md Phase C
rules: COM-confirmed correct on ≥3 contexts + minimal repro.
"""
import os
import zipfile

OUT = os.path.dirname(os.path.abspath(__file__))
LATIN = "Times New Roman"
CJK = "ＭＳ 明朝"


def run(text, sz=24, strike=False, dstrike=False, bold=False, color=None):
    extras = ""
    if strike:
        extras += "<w:strike/>"
    if dstrike:
        extras += "<w:dstrike/>"
    if bold:
        extras += "<w:b/>"
    if color:
        extras += f'<w:color w:val="{color}"/>'
    return (
        "<w:r><w:rPr>"
        f'<w:rFonts w:ascii="{LATIN}" w:hAnsi="{LATIN}" w:eastAsia="{CJK}" w:hint="eastAsia"/>'
        f"{extras}"
        f'<w:sz w:val="{sz}"/>'
        "</w:rPr>"
        f"<w:t xml:space=\"preserve\">{text}</w:t></w:r>"
    )


def para(runs_xml):
    return (
        "<w:p><w:pPr><w:jc w:val=\"left\"/>"
        '<w:spacing w:before="0" w:after="120" w:line="360" w:lineRule="auto"/>'
        "</w:pPr>"
        f"{runs_xml}</w:p>"
    )


def make_doc(paras_xml, page_w_tw=11906, margin_tw=1418):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>"
        f"{paras_xml}"
        f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
        f'<w:pgMar w:top="1418" w:right="{margin_tw}" w:bottom="1418" w:left="{margin_tw}"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:linePitch="360"/>'
        "</w:sectPr></w:body></w:document>"
    )


def make_styles(sz=24):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:docDefaults><w:rPrDefault><w:rPr>"
        f'<w:rFonts w:ascii="{LATIN}" w:hAnsi="{LATIN}" w:eastAsia="{CJK}" w:hint="eastAsia"/>'
        f'<w:sz w:val="{sz}"/>'
        "</w:rPr></w:rPrDefault>"
        "<w:pPrDefault><w:pPr/></w:pPrDefault>"
        "</w:docDefaults></w:styles>"
    )


def make_settings():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )


def write_docx(path, doc_xml):
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        "</Types>"
    )
    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        "</Relationships>"
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
        ' Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
        ' Target="settings.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", make_styles())
        z.writestr("word/settings.xml", make_settings())


def main():
    # DS_full: 12pt variants — typical body text size.
    paras = "".join([
        para(run("L1 plain — no decoration (control)")),
        para(run("L2 single strike — w:strike", strike=True)),
        para(run("L3 double strike — w:dstrike (R66 target)", dstrike=True)),
        para(run("L4 dstrike + bold", dstrike=True, bold=True)),
        para(run("L5 dstrike CJK — 二重打消し線テスト", dstrike=True)),
        para(run("L6 dstrike red", dstrike=True, color="FF0000")),
    ])
    write_docx(os.path.join(OUT, "DS_full.docx"), make_doc(paras))
    print(f"wrote {os.path.join(OUT, 'DS_full.docx')}")

    # DS_large: 24pt — gap is easier to see visually.
    large_paras = "".join([
        para(run("plain", sz=48)),
        para(run("strike", sz=48, strike=True)),
        para(run("DSTRIKE", sz=48, dstrike=True)),
    ])
    write_docx(os.path.join(OUT, "DS_large.docx"), make_doc(large_paras))
    print(f"wrote {os.path.join(OUT, 'DS_large.docx')}")


if __name__ == "__main__":
    main()
