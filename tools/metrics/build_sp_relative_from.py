"""§17.3 expansion (round 2): test positionV `relativeFrom` reference values.

Test: positionV with relativeFrom ∈ {paragraph, page, margin, line, topMargin, bottomMargin}
× posOffset ∈ {0, 100pt = 1270000 EMU} (just enough to discriminate references).

Per ECMA-376 ST_RelFromV, the valid values are:
  paragraph, page, margin, line, topMargin, bottomMargin, character,
  insidemargin, outsidemargin

Already confirmed paragraph (12 variants). Now test the rest.
"""
import zipfile
from pathlib import Path

OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\sp_relfrom_repro")
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


def shape_para(label: str, rel_v: str, posV_emu: int, rel_h: str = "column", posH_emu: int = 0) -> str:
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0"'
        ' w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
        ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="21"/></w:rPr>'
        f'<w:t>{label}</w:t>'
        '<mc:AlternateContent><mc:Choice Requires="wps">'
        '<w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0"'
        ' simplePos="0" relativeHeight="1" behindDoc="0" locked="0"'
        ' layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="{rel_h}"><wp:posOffset>{posH_emu}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="{rel_v}"><wp:posOffset>{posV_emu}</wp:posOffset></wp:positionV>'
        '<wp:extent cx="1270000" cy="762000"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapNone/>'
        f'<wp:docPr id="1" name="rect"/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr/>'
        '<wps:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="762000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>'
        '</wps:spPr><wps:bodyPr rot="0"/></wps:wsp>'
        '</a:graphicData></a:graphic>'
        '</wp:anchor></w:drawing>'
        '</mc:Choice></mc:AlternateContent>'
        '</w:r></w:p>'
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


def body_with_shape(rel_v, posV_emu):
    return (
        body_para("body para A")
        + body_para("body para B")
        + shape_para("anchor para C with shape", rel_v, posV_emu)
        + body_para("tail paragraph")
    )


# Test relativeFrom variants × posOffset ∈ {0, 1270000=100pt}
REL_FROMS = ["paragraph", "page", "margin", "line", "topMargin", "bottomMargin"]
POS_OFFSETS = [0, 1270000]   # 0pt, 100pt

CASES = []
for rv in REL_FROMS:
    for po in POS_OFFSETS:
        po_pt = po / 12700.0
        CASES.append((f"SR_{rv}_pos{int(po_pt)}", rv, po))


def main():
    for name, rv, po in CASES:
        make_docx(name, body_with_shape(rv, po))
    print(f"\n{len(CASES)} variants written")


if __name__ == "__main__":
    main()
