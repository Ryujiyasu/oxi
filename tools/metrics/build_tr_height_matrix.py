"""§13.5 + §19.4 cross-validation: trHeight behavior matrix.

§19.4 (DERIVED) claim: "Word does NOT grid-snap when trHeight is present"
  — based on 2ea81a tbl#1 atLeast=510tw=25.5pt 2-line cells observed
    row height = 25.5pt (specified value, no grid-snap effect).

§13.5 (CONFIRMED Round 22, 2026-04-08) rule:
  auto    -> content_height (grid-snapped per docGrid)
  exact   -> specified (regardless of content)
  atLeast -> max(content_height, specified)

Question: in atLeast mode where content > specified (content wins), is
content grid-snapped? §13.5 Round 22 data suggested yes (atLeast=25, 3
lines → 54.5 ≈ 3 × 18pt grid). §19.4 said no for 2ea81a. Discrepancy.

Build TR_* matrix:
  rule ∈ {auto, atLeast, exact}
  trHeight_tw ∈ {0, 200, 400, 510, 800, 1200}  (0..60pt)
  content_lines ∈ {1, 2, 3}    (lines of MS Mincho 10.5pt content)
  docGrid linePitch ∈ {323, 360}  (16.15pt and 18pt grid)

That's 3 × 6 × 3 × 2 = 108. Trim to a minimal informative subset:

Naming: TR_<rule>_h<tw>_<lines>L_lp<linePitch>.docx
"""
import zipfile
from pathlib import Path

OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\tr_height_repro")
OUT_DIR.mkdir(parents=True, exist_ok=True)


HEADER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body>'
)


def sect(line_pitch: int) -> str:
    return (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
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


def cell_paras(n_lines: int) -> str:
    """n_lines body paragraphs of Japanese text inside a cell."""
    return ''.join(body_para(f"行{i+1}テキスト") for i in range(n_lines))


def inline_tbl(rule: str, tw: int, n_lines: int) -> str:
    """1-row 1-col table with controlled trHeight + content."""
    if rule == "auto":
        # No w:hRule => default "auto"
        trh = f'<w:trHeight w:val="{tw}"/>' if tw > 0 else ""
    else:
        trh = f'<w:trHeight w:val="{tw}" w:hRule="{rule}"/>'
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
        f'<w:tr><w:trPr>{trh}</w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9638"/></w:tcPr>'
        f'{cell_paras(n_lines)}'
        '</w:tc></w:tr></w:tbl>'
    )


def make_docx(name: str, body_xml: str, line_pitch: int):
    out = OUT_DIR / f"{name}.docx"
    doc_xml = HEADER + body_xml + sect(line_pitch)
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


def body(rule, tw, lines):
    return body_para("anchor body para") + inline_tbl(rule, tw, lines) + body_para("tail paragraph")


def main():
    cases = []
    rules = ["auto", "atLeast", "exact"]
    tw_values = [200, 400, 510, 800, 1200]   # 10pt, 20pt, 25.5pt, 40pt, 60pt
    line_counts = [1, 2, 3]
    line_pitches = [323, 360]                  # 16.15pt and 18pt grids

    for rule in rules:
        for tw in tw_values:
            for lines in line_counts:
                for lp in line_pitches:
                    name = f"TR_{rule}_h{tw}_{lines}L_lp{lp}"
                    cases.append((name, rule, tw, lines, lp))

    for name, rule, tw, lines, lp in cases:
        b = body(rule, tw, lines)
        make_docx(name, b, line_pitch=lp)
    print(f"\n{len(cases)} variants")


if __name__ == "__main__":
    main()
