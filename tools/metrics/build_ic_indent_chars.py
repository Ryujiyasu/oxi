"""§15.1 expansion: which font's size determines char_width for
`<w:ind w:leftChars="N"/>` resolution?

Spec §15.1 says: `effective_indent = leftChars / 100 × char_width`.
For CJK fullwidth, char_width = font_size_pt. But which font?
- docDefault.rPr.sz ?
- Paragraph pPr.rPr.sz ?
- Run rPr.sz ?

Build IC_* matrix isolating each axis. Each repro has a single body
paragraph with leftChars=100 (= 1 character wide). Measured indent
divided by 1 char_width should reveal which font wins.

Cases:
  IC_runOnly_sz21:    only run sz=21 (10.5pt)
  IC_runOnly_sz28:    only run sz=28 (14pt)
  IC_pprOnly_sz21:    only pPr.rPr sz=21
  IC_pprOnly_sz28:    only pPr.rPr sz=28
  IC_pprAndRun_p21r28: pPr.rPr sz=21, run sz=28 (which wins?)
  IC_pprAndRun_p28r21: pPr.rPr sz=28, run sz=21 (which wins?)
  IC_runOnly_sz21_emptyRun: empty <w:r/> (no run text) with sz=21
"""
import zipfile
from pathlib import Path

OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\ic_repro")
OUT_DIR.mkdir(parents=True, exist_ok=True)


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


def para(left_chars: int, ppr_sz: int = 0, run_sz: int = 0,
         empty_run: bool = False, text: str = "サンプル文字列") -> str:
    ppr = '<w:pPr>'
    ppr += f'<w:ind w:leftChars="{left_chars}"/>'
    if ppr_sz:
        ppr += (
            '<w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
            ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/>'
            f'<w:sz w:val="{ppr_sz}"/>'
            '</w:rPr>'
        )
    ppr += '</w:pPr>'

    if empty_run:
        return f'<w:p>{ppr}</w:p>'

    run = '<w:r>'
    if run_sz:
        run += (
            '<w:rPr>'
            '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
            ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/>'
            f'<w:sz w:val="{run_sz}"/>'
            '</w:rPr>'
        )
    run += f'<w:t>{text}</w:t></w:r>'

    return f'<w:p>{ppr}{run}</w:p>'


def make_docx(name: str, body_xml: str):
    out = OUT_DIR / f"{name}.docx"
    doc_xml = HEADER + body_xml + FOOTER
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
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
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
    print(f"Wrote {out.name}")


# Variants
CASES = [
    # (name, ppr_sz, run_sz, empty_run)
    ("IC_runOnly_sz21",        0, 21, False),
    ("IC_runOnly_sz28",        0, 28, False),
    ("IC_pprOnly_sz21",       21,  0, False),
    ("IC_pprOnly_sz28",       28,  0, False),
    ("IC_pprAndRun_p21_r28", 21, 28, False),
    ("IC_pprAndRun_p28_r21", 28, 21, False),
    ("IC_emptyRun_p21",       21,  0, True),
    ("IC_emptyRun_p28",       28,  0, True),
    ("IC_emptyRun_none",       0,  0, True),  # baseline default
    # Mixed run with first run differing from later run
]


def main():
    for name, ppr_sz, run_sz, empty in CASES:
        body = para(left_chars=100, ppr_sz=ppr_sz, run_sz=run_sz, empty_run=empty)
        # Append a tail so layout is well-formed
        body += '<w:p><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>tail paragraph</w:t></w:r></w:p>'
        make_docx(name, body)
    print(f"\n{len(CASES)} variants written")


if __name__ == "__main__":
    main()
