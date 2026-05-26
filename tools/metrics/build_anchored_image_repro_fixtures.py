"""Author minimal `<wp:anchor>` (floating / anchored image) repro
fixtures for S316.

`image_integration.rs` (S292) covers INLINE images
(`<wp:inline>`) — extent EMU→pt, alt_text, embedded PNG blob. But
`parse_drawing` at parser/ooxml.rs:2958 has a LARGER surface for
anchored images (`<wp:anchor>`):
  - positionH / positionV with posOffset (EMU/12700 → pt) OR align
    (string)
  - relativeFrom attr → h_relative / v_relative (string enum)
  - wrap{None,Square,Tight,TopAndBottom} → WrapType enum
  - srcRect l/t/r/b → ImageCrop (val/1000 → percent)
  - Routing: image.position.is_some() → Page.floating_images
    (NOT Block::Image in paragraph). This is the structural diff
    from inline.

Outputs to ``tools/fixtures/anchored_image_samples/``.
"""
import os
import struct
import zlib
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "anchored_image_samples")


def minimal_png(w: int = 1, h: int = 1, rgba=(255, 255, 255, 255)) -> bytes:
    r, g, b, a = rgba
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr_data = struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0)
    ihdr_chunk = b"IHDR" + ihdr_data
    ihdr = (
        struct.pack(">I", len(ihdr_data))
        + ihdr_chunk
        + struct.pack(">I", zlib.crc32(ihdr_chunk))
    )
    raw = b""
    for _ in range(h):
        raw += bytes([0]) + bytes([r, g, b, a]) * w
    compressed = zlib.compress(raw)
    idat_chunk = b"IDAT" + compressed
    idat = (
        struct.pack(">I", len(compressed))
        + idat_chunk
        + struct.pack(">I", zlib.crc32(idat_chunk))
    )
    iend_chunk = b"IEND"
    iend = (
        struct.pack(">I", 0)
        + iend_chunk
        + struct.pack(">I", zlib.crc32(iend_chunk))
    )
    return sig + ihdr + idat + iend


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="ＭＳ 明朝" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
</w:styles>"""

DOC_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<w:body>
"""

SECT_PR = (
    "<w:sectPr>"
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    "</w:sectPr>"
)


def _anchor_drawing(
    extent_cx: int = 914400,
    extent_cy: int = 914400,
    pos_h: str = "",
    pos_v: str = "",
    wrap: str = "",
    src_rect: str = "",
    descr: str | None = None,
) -> str:
    descr_attr = f' descr="{escape(descr)}"' if descr is not None else ""
    src_rect_xml = src_rect if src_rect else ""
    return f"""<w:r><w:drawing>
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
{pos_h}
{pos_v}
<wp:extent cx="{extent_cx}" cy="{extent_cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
{wrap}
<wp:docPr id="1" name="Anchored"{descr_attr}/>
<wp:cNvGraphicFramePr/>
<a:graphic>
<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic>
<pic:nvPicPr>
<pic:cNvPr id="1" name="Anchored"/>
<pic:cNvPicPr/>
</pic:nvPicPr>
<pic:blipFill>
<a:blip r:embed="rIdImg"/>
{src_rect_xml}
<a:stretch><a:fillRect/></a:stretch>
</pic:blipFill>
<pic:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="{extent_cx}" cy="{extent_cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</pic:spPr>
</pic:pic>
</a:graphicData>
</a:graphic>
</wp:anchor>
</w:drawing></w:r>"""


def _paragraph(*runs: str) -> str:
    return f'<w:p>{"".join(runs)}</w:p>'


def write_docx(path: str, body_xml: str, png_bytes: bytes) -> None:
    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
        z.writestr("word/media/image1.png", png_bytes)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")
    png = minimal_png()

    # v1_anchor_pos_offset: posOffset for both H and V. Pins:
    #   - posOffset Text content → EMU/12700 → pt
    #     (914400 / 12700 = 72.0pt; 457200 / 12700 = 36.0pt)
    #   - relativeFrom attr captured as h_relative / v_relative
    pos_h = """<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>"""
    pos_v = """<wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>"""
    body = _paragraph(_anchor_drawing(pos_h=pos_h, pos_v=pos_v,
                                       wrap="<wp:wrapNone/>",
                                       descr="positioned"))
    write_docx(os.path.join(OUT_DIR, "v1_anchor_pos_offset.docx"), body, png)

    # v1_anchor_pos_align: align string for both H and V. Pins the
    # alternative position branch (align Text → h_align/v_align
    # string) and that pos_x/pos_y stay at 0.0 default.
    pos_h = """<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>"""
    pos_v = """<wp:positionV relativeFrom="page"><wp:align>top</wp:align></wp:positionV>"""
    body = _paragraph(_anchor_drawing(pos_h=pos_h, pos_v=pos_v,
                                       wrap="<wp:wrapNone/>",
                                       descr="aligned"))
    write_docx(os.path.join(OUT_DIR, "v1_anchor_pos_align.docx"), body, png)

    # v1_anchor_wrap_square: <wp:wrapSquare/> → WrapType::Square.
    pos_h = """<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>"""
    pos_v = """<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>"""
    body = _paragraph(_anchor_drawing(
        pos_h=pos_h, pos_v=pos_v,
        wrap='<wp:wrapSquare wrapText="bothSides"/>',
        descr="square-wrap",
    ))
    write_docx(os.path.join(OUT_DIR, "v1_anchor_wrap_square.docx"), body, png)

    # v1_anchor_wrap_topandbottom: <wp:wrapTopAndBottom/> →
    # WrapType::TopAndBottom. Distinct enum from Square / None / Tight.
    body = _paragraph(_anchor_drawing(
        pos_h=pos_h, pos_v=pos_v,
        wrap="<wp:wrapTopAndBottom/>",
        descr="top-and-bottom-wrap",
    ))
    write_docx(os.path.join(OUT_DIR, "v1_anchor_wrap_topandbottom.docx"),
               body, png)

    # v1_anchor_crop_srcrect: srcRect l/t/r/b in 1/1000th percent units.
    # Parser at ooxml.rs:3457 divides by 1000 → percent.
    #   l=10000 → 10.0%
    #   t=20000 → 20.0%
    #   r=30000 → 30.0%
    #   b=40000 → 40.0%
    src_rect = '<a:srcRect l="10000" t="20000" r="30000" b="40000"/>'
    body = _paragraph(_anchor_drawing(
        pos_h=pos_h, pos_v=pos_v,
        wrap="<wp:wrapNone/>",
        src_rect=src_rect,
        descr="cropped",
    ))
    write_docx(os.path.join(OUT_DIR, "v1_anchor_crop_srcrect.docx"),
               body, png)

    print("Done.")


if __name__ == "__main__":
    main()
