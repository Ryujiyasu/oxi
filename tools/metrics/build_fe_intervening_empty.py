"""§19.7 follow-up: investigate the +28.55pt Y0 intercept anomaly observed
in 2ea81a tbl#3 (3 intervening empty body paras between tbl#2 and tbl#3,
Y0 = anchor_top + 28.55pt = 2 × 14.25pt = 2 line heights of MS Mincho 12pt).

Hypothesis A: Each intervening empty paragraph adds +1 line height to the
              Y0 intercept (+anchor_top reported by COM is the LAST empty
              para, but Word internally references the body-flow position
              N line heights earlier).

Hypothesis B: The anomaly only appears when a floating table precedes the
              empty paragraphs. With only body-para precedent, intercept is
              still 1 line height regardless of empty para count.

Hypothesis C: The anomaly is a 2ea81a-specific structural quirk (e.g.,
              tblpY value, font size, line pitch interaction).

Repro matrix (each at tblpY = 0 and 600tw):
  FE_para_0e    : 1 body para + 0 empty + tbl
  FE_para_1e    : 1 body para + 1 empty + tbl
  FE_para_2e    : 1 body para + 2 empty + tbl
  FE_para_3e    : 1 body para + 3 empty + tbl   (matches 2ea81a empty count)
  FE_para_5e    : 1 body para + 5 empty + tbl
  FE_inlinetbl_3e : body para + inline tbl + 3 empty + tbl  (table-precedent)
  FE_floattbl_3e  : body para + floating tbl A + 3 empty + floating tbl B
                    (closest to 2ea81a tbl#2 -> tbl#3 structure)

Output: tools/metrics/fe_repro/FE_*.docx
"""
import zipfile
from pathlib import Path

OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\fe_repro")
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


def empty_para() -> str:
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0"'
        ' w:line="240" w:lineRule="auto"/>'
        '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
        ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="21"/></w:rPr>'
        '</w:pPr></w:p>'
    )


def inline_tbl(text: str) -> str:
    return (
        '<w:tbl><w:tblPr>'
        '<w:tblW w:type="dxa" w:w="9638"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9638"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9638"/></w:tcPr>'
        f'{body_para(text)}'
        '</w:tc></w:tr></w:tbl>'
    )


def floating_tbl(tblpY_tw: int, color: str, label: str) -> str:
    return (
        '<w:tbl><w:tblPr>'
        f'<w:tblpPr w:leftFromText="142" w:rightFromText="142"'
        f' w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="{tblpY_tw}"/>'
        '<w:tblW w:type="dxa" w:w="9638"/>'
        '<w:tblBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9638"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9638"/></w:tcPr>'
        f'{body_para(label)}'
        '</w:tc></w:tr></w:tbl>'
    )


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


# Variant matrix
def precedent_para_only(n_empty: int) -> str:
    """1 body para + n empty paras."""
    return body_para("anchor para A") + (empty_para() * n_empty)


def precedent_inline_tbl(n_empty: int) -> str:
    """1 body para + 1 inline tbl + n empty paras."""
    return body_para("body para A") + inline_tbl("inline cell") + (empty_para() * n_empty)


def precedent_floating_tbl(n_empty: int, tblpY_outer_tw: int) -> str:
    """1 body para + 1 floating tbl A + n empty paras (the trailing floating tbl
    will be appended by the caller). The outer floating tbl varies tblpY too,
    but we fix it at one value to isolate empty-para effect."""
    return (
        body_para("body para A")
        + floating_tbl(tblpY_outer_tw, "008000", "outer floating tbl A")
        + (empty_para() * n_empty)
    )


PRE_KINDS = [
    ("para",     precedent_para_only,    [0, 1, 2, 3, 5]),
    ("inlinetbl", precedent_inline_tbl,   [0, 1, 2, 3, 5]),
    # For floating-table precedent we vary at one fixed outer tblpY=225tw
    # (matches 2ea81a's tbl#2 setting); empty count varied.
    ("floattbl225", lambda n: precedent_floating_tbl(n, 225), [0, 1, 2, 3, 5]),
]
TBLPY_VALUES = [("Y0", 0), ("Y600", 600)]


def main():
    n = 0
    for pkname, prefn, empty_counts in PRE_KINDS:
        for ec in empty_counts:
            for yname, ytw in TBLPY_VALUES:
                pre = prefn(ec)
                inner = floating_tbl(ytw, "0000FF", f"target floating {pkname} {ec}e {yname}")
                tail = body_para("tail paragraph")
                body = pre + inner + tail
                name = f"FE_{pkname}_{ec}e_{yname}"
                make_docx(name, body)
                n += 1
    print(f"\n{n} variants written to {OUT_DIR}")


if __name__ == "__main__":
    main()
