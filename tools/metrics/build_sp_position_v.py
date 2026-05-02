"""§17.1 expansion: pin formula for `<wp:positionV relativeFrom="paragraph">`
+ <wp:posOffset>VAL</wp:posOffset> (VAL in EMU; pt = VAL / 12700).

Hypothesis: shape_top_pt = anchor_paragraph_top_pt + posOffset_pt.

Per §17.1, "v_rel=2 (paragraph)" reference is the "anchor paragraph",
but no formula or shape-top vs paragraph-top semantics is defined.

Build SP_* test grid:
  axis 1: anchor position
    - first body paragraph
    - 3rd body paragraph (after 2 prior)
    - paragraph inside a single-cell table
  axis 2: posOffset (EMU)
    - 0  (0 pt)
    - 114300  (9 pt — most common in baseline)
    - 675640  (53.2 pt)
    - -635000 (-50 pt — negative test)
  axis 3: wrap mode
    - wrapNone (172/177 in baseline)

That's 3 × 4 × 1 = 12 SP_* variants.
"""
import zipfile
from pathlib import Path

OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\sp_repro")
OUT_DIR.mkdir(parents=True, exist_ok=True)

HEADER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
    ' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"'
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' mc:Ignorable="wp14">'
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


def shape_paragraph(label: str, posOffsetV_emu: int, posOffsetH_emu: int = 0) -> str:
    """A body paragraph that contains a wp:anchor wrapping a rectangle shape.
    Shape is fixed 100×60 pt (= 1270000 × 762000 EMU)."""
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0"'
        ' w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
        ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="21"/></w:rPr>'
        f'<w:t>{label}</w:t>'
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="wps">'
        '<w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0"'
        ' simplePos="0" relativeHeight="1" behindDoc="0" locked="0"'
        ' layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:posOffset>{posOffsetH_emu}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{posOffsetV_emu}</wp:posOffset></wp:positionV>'
        '<wp:extent cx="1270000" cy="762000"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapNone/>'
        f'<wp:docPr id="1" name="rect_{abs(posOffsetV_emu)}"/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp>'
        '<wps:cNvSpPr/>'
        '<wps:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="762000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>'
        '</wps:spPr>'
        '<wps:bodyPr rot="0"/>'
        '</wps:wsp>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:anchor>'
        '</w:drawing>'
        '</mc:Choice>'
        '</mc:AlternateContent>'
        '</w:r></w:p>'
    )


def cell_paragraph_with_shape(label: str, posOffsetV_emu: int) -> str:
    """A 1×1 inline table whose cell paragraph contains the wp:anchor."""
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
        f'{shape_paragraph(label, posOffsetV_emu)}'
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


# Anchor variants
def anchor_first(posV_emu):
    return (
        shape_paragraph(f"anchor para A (shape host)", posV_emu)
        + body_para("body para B")
        + body_para("body para C")
    )


def anchor_third(posV_emu):
    return (
        body_para("body para A")
        + body_para("body para B")
        + shape_paragraph(f"anchor para C (shape host)", posV_emu)
    )


def anchor_in_cell(posV_emu):
    return (
        body_para("body before tbl")
        + cell_paragraph_with_shape("cell-anchor (shape)", posV_emu)
        + body_para("body after tbl")
    )


CASES = [
    ("SP_first_pos0",         anchor_first,    0),
    ("SP_first_pos9pt",       anchor_first,    114300),
    ("SP_first_pos53pt",      anchor_first,    675640),
    ("SP_first_neg50pt",      anchor_first,   -635000),
    ("SP_third_pos0",         anchor_third,    0),
    ("SP_third_pos9pt",       anchor_third,    114300),
    ("SP_third_pos53pt",      anchor_third,    675640),
    ("SP_third_neg50pt",      anchor_third,   -635000),
    ("SP_cell_pos0",          anchor_in_cell,  0),
    ("SP_cell_pos9pt",        anchor_in_cell,  114300),
    ("SP_cell_pos53pt",       anchor_in_cell,  675640),
    ("SP_cell_neg50pt",       anchor_in_cell, -635000),
]


def main():
    for name, fn, posV in CASES:
        body = fn(posV)
        # Append a tail body para to ensure layout doesn't truncate
        body += body_para("tail paragraph")
        make_docx(name, body)
    print(f"\n{len(CASES)} variants written")


if __name__ == "__main__":
    main()
