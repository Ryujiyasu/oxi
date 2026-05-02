"""§15.1 follow-up: with explicit docDefaults, does char_width for
leftChars depend on docDefault.rPr.sz?
"""
import zipfile
from pathlib import Path

OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\ic_docdef_repro")
OUT_DIR.mkdir(parents=True, exist_ok=True)


def make_styles_xml(default_sz: int) -> str:
    """styles.xml with docDefaults setting <w:sz w:val="default_sz"/>."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults>'
        '<w:rPrDefault>'
        '<w:rPr>'
        '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
        ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/>'
        f'<w:sz w:val="{default_sz}"/>'
        '</w:rPr>'
        '</w:rPrDefault>'
        '<w:pPrDefault><w:pPr/></w:pPrDefault>'
        '</w:docDefaults>'
        '</w:styles>'
    )


HEADER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body>'
)
FOOTER = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"'
    ' w:header="720" w:footer="720" w:gutter="0"/>'
    '<w:cols w:space="720"/>'
    '<w:docGrid w:linePitch="360"/>'
    '</w:sectPr>'
    '</w:body></w:document>'
)


def para(left_chars: int, run_sz: int = 0) -> str:
    if run_sz:
        run = (
            f'<w:r><w:rPr><w:sz w:val="{run_sz}"/></w:rPr>'
            '<w:t>サンプル</w:t></w:r>'
        )
    else:
        run = '<w:r><w:t>サンプル</w:t></w:r>'
    return f'<w:p><w:pPr><w:ind w:leftChars="{left_chars}"/></w:pPr>{run}</w:p>'


def make_docx(name: str, body_xml: str, default_sz: int):
    out = OUT_DIR / f"{name}.docx"
    doc_xml = HEADER + body_xml + FOOTER
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>'
    )
    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>'
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
        ' Target="styles.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", make_styles_xml(default_sz))
    print(f"Wrote {out.name} (default_sz={default_sz})")


# Default sizes to test
CASES = [
    ("IC_dd_sz20",       20),  # 10pt
    ("IC_dd_sz21",       21),  # 10.5pt — Word default
    ("IC_dd_sz24",       24),  # 12pt
    ("IC_dd_sz28",       28),  # 14pt
    ("IC_dd_sz44",       44),  # 22pt
]


def main():
    for name, default_sz in CASES:
        body = para(left_chars=100)
        body += '<w:p><w:r><w:t>tail</w:t></w:r></w:p>'
        make_docx(name, body, default_sz)
    # Also build runOverride: docDef sz=20 BUT run sz=44
    body = para(left_chars=100, run_sz=44)
    body += '<w:p><w:r><w:t>tail</w:t></w:r></w:p>'
    make_docx("IC_dd20_run44", body, 20)

    body = para(left_chars=100, run_sz=21)
    body += '<w:p><w:r><w:t>tail</w:t></w:r></w:p>'
    make_docx("IC_dd44_run21", body, 44)
    print()


if __name__ == "__main__":
    main()
