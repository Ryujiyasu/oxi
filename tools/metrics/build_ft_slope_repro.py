"""Build FT_* (Floating Table) minimal repros to identify when Word
respects vs ignores tblpY.

Background:
  Existing TP1-6 repros showed table_top INVARIANT to tblpY (slope=0):
    TP1 tblpY=2.5pt  -> table_top=71.0
    TP2 tblpY=10pt   -> table_top=71.0
    TP3 tblpY=30pt   -> table_top=71.0
  (anchor para "anchor para line 1" at y=56.5pt, line_height ~14pt)

  But 2ea81a tblpY sweep on tbl#2 showed slope=1.0:
    V1_zero    tblpY=0   -> 610.5
    V0_orig    tblpY=11.25 -> 621.5
    V4_1000tw  tblpY=50  -> 660.5
    V5_2000tw  tblpY=100 -> 710.5

  Difference unknown. This script builds repros varying:
    - PRE_KIND: what immediately precedes the floating table
        (body_para, inline_tbl, page_break, multi_para)
    - TBLPY: small (50tw=2.5pt), medium (600tw=30pt), large (4000tw=200pt)

Output: tools/metrics/ft_slope_repro/FT_*.docx

Naming: FT_<PreKind>_Y<tw>.docx
"""
import shutil
import zipfile
from pathlib import Path

OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\ft_slope_repro")
OUT_DIR.mkdir(parents=True, exist_ok=True)


HEADER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
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


def body_para(text: str) -> str:
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0"'
        ' w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
        ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="21"/></w:rPr>'
        f'<w:t>{text}</w:t></w:r></w:p>'
    )


def empty_body_para() -> str:
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0"'
        ' w:line="240" w:lineRule="auto"/></w:pPr></w:p>'
    )


def inline_tbl(text: str) -> str:
    """Inline (non-floating) 1x1 table."""
    return (
        '<w:tbl><w:tblPr>'
        '<w:tblW w:type="dxa" w:w="9638"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9638"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9638"/></w:tcPr>'
        f'{body_para(text)}'
        '</w:tc></w:tr></w:tbl>'
    )


def floating_tbl(tblpY_tw: int, label: str) -> str:
    """Floating 1x1 table with vertAnchor=text, horzAnchor=margin."""
    return (
        '<w:tbl><w:tblPr>'
        f'<w:tblpPr w:leftFromText="142" w:rightFromText="142"'
        f' w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="{tblpY_tw}"/>'
        '<w:tblW w:type="dxa" w:w="9638"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="0000FF"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="0000FF"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="0000FF"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="0000FF"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9638"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9638"/></w:tcPr>'
        f'{body_para(label)}'
        '</w:tc></w:tr></w:tbl>'
    )


def make_docx(name: str, body_xml: str):
    """Write minimal .docx with given body XML."""
    doc_xml = HEADER + body_xml + FOOTER
    out = OUT_DIR / f"{name}.docx"

    # Minimal package: [Content_Types].xml, _rels/.rels, word/_rels/document.xml.rels,
    # word/document.xml, word/styles.xml
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


# Variant matrix: PRE_KIND × tblpY
# PRE_KIND values:
#   1para    : single body paragraph
#   3para    : three body paragraphs
#   1empty   : single empty paragraph
#   inline   : single inline (non-floating) table
#   inline_p : inline table followed by 1 body paragraph
TBLPY_VALUES = [
    ("Y0",     0),     # 0pt
    ("Y50",   50),     # 2.5pt
    ("Y600", 600),     # 30pt
    ("Y2000", 2000),   # 100pt
    ("Y4000", 4000),   # 200pt
]

PRE_KINDS = {
    "1para": lambda: body_para("anchor para A"),
    "3para": lambda: (
        body_para("para A line 1") +
        body_para("para A line 2") +
        body_para("para A line 3")
    ),
    "1empty": lambda: empty_body_para(),
    "inline": lambda: inline_tbl("inline table content"),
    "inline_p": lambda: inline_tbl("inline table content") + body_para("after-inline para"),
}

VARIANTS = []
for pkname, prefn in PRE_KINDS.items():
    for yname, ytw in TBLPY_VALUES:
        VARIANTS.append((f"FT_{pkname}_{yname}", prefn(), ytw))


def main():
    for name, pre, ytw in VARIANTS:
        body = pre + floating_tbl(ytw, f"floating tbl {name} tblpY={ytw}tw") + body_para("tail paragraph")
        make_docx(name, body)
    print(f"\n{len(VARIANTS)} variants written to {OUT_DIR}")


if __name__ == "__main__":
    main()
