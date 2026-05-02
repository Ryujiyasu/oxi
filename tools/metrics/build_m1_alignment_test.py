"""§4.7 Mech 1 — alignment dependency test.

§4.7b confirmed Mech 2 fires only at jc=both. §4.7 (Mech 1 Type A/B/C) does
NOT specify alignment requirement. Hypothesis: Mech 1 is alignment-agnostic
(fires at any jc, gated only by `<w:kern>` per session 51).

Test design (each fixture = 1 doc with 5 paragraphs):
  Each paragraph contains: 漢）（漢 (4 chars; ）（ is B→A Mech 1 trigger)
  Paragraphs differ only in <w:jc> alignment value:
    p1: <w:jc w:val="both"/>     (justify)
    p2: <w:jc w:val="left"/>     (left)
    p3: <w:jc w:val="center"/>   (center)
    p4: <w:jc w:val="right"/>    (right)
    p5: (no <w:jc/>)             (default, typically "left" or inherited)

Two fixtures:
  M1A_kern_on.docx:  docDefaults includes <w:kern w:val="2"/>
  M1A_kern_off.docx: docDefaults has NO <w:kern> element

Page width 400pt, content 230pt — natural line width = 4×10.5 = 42pt
(or 21+5.5+5.5+10.5 = ~42 if compressed). Plenty of room → no overflow,
Mech 2 cannot fire. Any compression → Mech 1.

Per-char advance measured via Information(5).
"""
import zipfile
from pathlib import Path

OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\m1_alignment_repro")
OUT_DIR.mkdir(parents=True, exist_ok=True)


def styles_xml(kern_on: bool) -> str:
    kern_xml = '<w:kern w:val="2"/>' if kern_on else ''
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults>'
        '<w:rPrDefault>'
        '<w:rPr>'
        '<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
        ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/>'
        '<w:sz w:val="21"/>'
        f'{kern_xml}'
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
    '<w:pgSz w:w="8000" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"'
    ' w:header="720" w:footer="720" w:gutter="0"/>'
    '<w:cols w:space="720"/>'
    '<w:docGrid w:linePitch="360"/>'
    '</w:sectPr>'
    '</w:body></w:document>'
)


def para(jc_val: str | None, text: str = "漢）（漢") -> str:
    """One paragraph with optional w:jc."""
    jc_xml = f'<w:jc w:val="{jc_val}"/>' if jc_val else ''
    ppr = f'<w:pPr>{jc_xml}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    run = (
        '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
        ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/>'
        '<w:sz w:val="21"/></w:rPr>'
        f'<w:t>{text}</w:t></w:r>'
    )
    return f'<w:p>{ppr}{run}</w:p>'


def make_docx(name: str, kern_on: bool):
    out = OUT_DIR / f"{name}.docx"
    paras = ''.join([
        para("both"),
        para("left"),
        para("center"),
        para("right"),
        para(None),
    ])
    doc_xml = HEADER + paras + FOOTER

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
        z.writestr("word/styles.xml", styles_xml(kern_on))
    print(f"Wrote {out.name} (kern={'ON' if kern_on else 'OFF'})")


def main():
    make_docx("M1A_kern_on", kern_on=True)
    make_docx("M1A_kern_off", kern_on=False)


if __name__ == "__main__":
    main()
