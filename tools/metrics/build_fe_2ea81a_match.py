"""§19.7 follow-up #2: replicate 2ea81a tbl#3 structural details to isolate
which factor causes the +28.55pt Y0 intercept.

2ea81a tbl#3 differs from my prior FE repros in:
  - docGrid: type="lines" linePitch="323"  (vs my "linesAndChars" lp=360)
  - Last (anchor) empty para: <w:spacing w:line="296" w:lineRule="atLeast"/>
                              + <w:rPr><w:sz w:val="28"/></w:rPr>  (14pt)
  - pgSz: w:code="9"
  - tbl has <w:tblStyle w:val="aa"/> + tblpPr (correct order)
  - tbl tblW="0" type="auto" (vs my dxa 9638)

We build 6 axis-isolation variants:
  K_baseline_lp360_auto_sz21:
      original FE_para_3e_Y0 control (lp=360, sz=21, line=auto)
  K_lp323:
      change linePitch 360 -> 323 only
  K_lp323_atLeast296:
      lp=323 + last empty para line=296 atLeast (sz=21)
  K_lp323_atLeast296_sz28:
      lp=323 + line=296 atLeast + last empty para sz=28
  K_full_match:
      lp=323 + atLeast + sz=28 + tblStyle "aa" + tblW=auto + page code=9
  K_only_sz28:
      lp=360 baseline but only the empty-para sz=28 (isolate sz)
"""
import zipfile
from pathlib import Path

OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\fe_match_repro")
OUT_DIR.mkdir(parents=True, exist_ok=True)


def header():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>'
    )


def sect(line_pitch: int, page_code: bool = False):
    code = ' w:code="9"' if page_code else ''
    return (
        '<w:sectPr>'
        f'<w:pgSz w:w="11906" w:h="16838"{code}/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"'
        ' w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        f'<w:docGrid w:type="lines" w:linePitch="{line_pitch}"/>'
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


def empty_para(line=240, line_rule="auto", sz=21) -> str:
    return (
        '<w:p><w:pPr>'
        f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="{line_rule}"/>'
        '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
        f' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="{sz}"/></w:rPr>'
        '</w:pPr></w:p>'
    )


def floating_tbl(tblpY_tw: int, label: str,
                 use_tblStyle=False, tblW_auto=False) -> str:
    style = '<w:tblStyle w:val="aa"/>' if use_tblStyle else ''
    tblw = ('<w:tblW w:w="0" w:type="auto"/>'
            if tblW_auto else '<w:tblW w:type="dxa" w:w="9638"/>')
    return (
        '<w:tbl><w:tblPr>'
        f'{style}'
        f'<w:tblpPr w:leftFromText="142" w:rightFromText="142"'
        f' w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="{tblpY_tw}"/>'
        f'{tblw}'
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


def make_styles_xml() -> str:
    """Provide a simple TableNormal + 'aa' style chain so tblStyle="aa" works."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal">'
        '<w:name w:val="Normal Table"/></w:style>'
        '<w:style w:type="table" w:styleId="aa">'
        '<w:name w:val="Table Grid"/>'
        '<w:basedOn w:val="TableNormal"/>'
        '<w:tblPr>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
        '</w:tblPr></w:style>'
        '</w:styles>'
    )


def make_docx(name: str, body_xml: str, line_pitch: int = 360,
              page_code: bool = False, with_styles: bool = False):
    out = OUT_DIR / f"{name}.docx"
    doc_xml = header() + body_xml + sect(line_pitch, page_code)

    if with_styles:
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
        doc_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
            ' Target="styles.xml"/>'
            '</Relationships>'
        )
    else:
        content_types = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '</Types>'
        )
        doc_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        )

    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>'
    )

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
        if with_styles:
            z.writestr("word/styles.xml", make_styles_xml())
    print(f"Wrote {out.name}")


# Build a body with 1 anchor para + 3 empty paras + floating tbl + tail.
# The 3rd empty para (anchor for the tbl) has tweakable params.
def body(anchor_line, anchor_rule, anchor_sz, tbl_kwargs):
    return (
        body_para("body para A")
        + empty_para()
        + empty_para()
        + empty_para(line=anchor_line, line_rule=anchor_rule, sz=anchor_sz)
        + floating_tbl(tblpY_tw=289, label="target tbl", **tbl_kwargs)
        + body_para("tail paragraph")
    )


def main():
    # Each variant runs at tblpY=289tw (the 2ea81a tbl#3 value, =14.45pt).
    cases = [
        ("K_baseline",            240, "auto",    21, {"use_tblStyle": False, "tblW_auto": False}, 360, False, False),
        ("K_lp323",               240, "auto",    21, {"use_tblStyle": False, "tblW_auto": False}, 323, False, False),
        ("K_lp323_atLeast296",    296, "atLeast", 21, {"use_tblStyle": False, "tblW_auto": False}, 323, False, False),
        ("K_lp323_atLeast296_sz28",296,"atLeast", 28, {"use_tblStyle": False, "tblW_auto": False}, 323, False, False),
        ("K_only_sz28",           240, "auto",    28, {"use_tblStyle": False, "tblW_auto": False}, 360, False, False),
        ("K_only_atLeast296",     296, "atLeast", 21, {"use_tblStyle": False, "tblW_auto": False}, 360, False, False),
        ("K_full_match",          296, "atLeast", 28, {"use_tblStyle": True,  "tblW_auto": True},  323, True, True),
        # tblStyle/tblW alone
        ("K_tblStyle_only",       240, "auto",    21, {"use_tblStyle": True,  "tblW_auto": False}, 360, False, True),
        ("K_tblWauto_only",       240, "auto",    21, {"use_tblStyle": False, "tblW_auto": True},  360, False, False),
    ]
    for name, al, ar, asz, tbl_kw, lp, pcode, with_styles in cases:
        b = body(al, ar, asz, tbl_kw)
        make_docx(name, b, line_pitch=lp, page_code=pcode, with_styles=with_styles)
    print(f"\n{len(cases)} variants written to {OUT_DIR}")


if __name__ == "__main__":
    main()
