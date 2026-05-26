"""Author minimal mc:AlternateContent + w:object (OLE) repro fixtures
for S320.

Two distinct parser entry points covered:

  - `parse_alternate_content` at parser/ooxml.rs:4017:
    mc:AlternateContent prefers <mc:Choice> (DrawingML) over
    <mc:Fallback> (VML legacy). Three fixtures pin the three
    states: Choice-only, Choice+Fallback (Choice wins),
    Fallback-only.
  - `parse_ole_object` at parser/ooxml.rs:3931: <w:object> with an
    embedded <v:shape><v:imagedata r:id="..."/></v:shape> for the
    OLE preview. The parsed result is an Image with
    alt_text="OLE Object" HARDCODED, plus width/height from the
    <v:shape style="..."> CSS-like attribute via parse_css_length.

Outputs to ``tools/fixtures/alternate_content_ole_samples/``.
"""
import os
import struct
import zlib
import zipfile

OUT_DIR = os.path.join(
    os.path.dirname(__file__), "..", "fixtures",
    "alternate_content_ole_samples",
)


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
<w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
</w:styles>"""

# Namespaces: mc=markup-compatibility, w=wordprocessingml,
# r=relationships, wp=drawing/wordprocessingDrawing, a=drawing/main,
# pic=drawing/picture, v=urn:schemas-microsoft-com:vml.
DOC_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:o="urn:schemas-microsoft-com:office:office">
<w:body>
"""

SECT_PR = (
    "<w:sectPr>"
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    "</w:sectPr>"
)


def _inline_drawing(cx: int, cy: int, descr: str) -> str:
    # NOTE: intentionally OMIT <a:prstGeom prst="rect"/> inside
    # <pic:spPr>. parse_drawing at parser/ooxml.rs:3304 treats
    # prstGeom as a shape_type signal — including it would cause
    # an Image AND a Shape to both surface in DrawingResult, which
    # contaminates the parse_alternate_content "image-only" test.
    # Real Word docs include prstGeom for images, but a minimal
    # blipFill-only image still parses correctly.
    return f"""<w:drawing>
<wp:inline distT="0" distB="0" distL="0" distR="0">
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:docPr id="1" name="img" descr="{descr}"/>
<wp:cNvGraphicFramePr/>
<a:graphic>
<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic>
<pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></pic:spPr>
</pic:pic>
</a:graphicData>
</a:graphic>
</wp:inline>
</w:drawing>"""


def _vml_pict_rect(width_pt: float, height_pt: float, fill: str) -> str:
    return (
        f'<w:pict>'
        f'<v:rect style="width:{width_pt}pt;height:{height_pt}pt" '
        f'fillcolor="#{fill}"></v:rect>'
        f'</w:pict>'
    )


def write_docx(path: str, body_xml: str, png_bytes: bytes,
               include_image_rel: bool = True) -> None:
    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    rels = DOC_RELS if include_image_rel else (
        """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>""")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
        if include_image_rel:
            z.writestr("word/media/image1.png", png_bytes)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")
    png = minimal_png()

    # ===== parse_alternate_content =====

    # v1_ac_choice_only_drawing: <mc:AlternateContent> with ONLY
    # <mc:Choice><w:drawing>. The parser must return DrawingML from
    # Choice (line 4036-4042). Image routes to inline (Block::Image).
    body = (
        '<w:p><w:r>'
        '<mc:AlternateContent>'
        '<mc:Choice Requires="wps">'
        + _inline_drawing(914400, 914400, "from-choice") +
        '</mc:Choice>'
        '</mc:AlternateContent>'
        '</w:r></w:p>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_ac_choice_only_drawing.docx"),
               body, png)

    # v1_ac_choice_wins_over_fallback: BOTH Choice (DrawingML) and
    # Fallback (VML pict). Parser prefers Choice (line 4036) and only
    # falls back to pict if result is still None (line 4043). Image
    # via Choice → Block::Image with alt_text="from-choice" (NOT a
    # VML shape from Fallback).
    body = (
        '<w:p><w:r>'
        '<mc:AlternateContent>'
        '<mc:Choice Requires="wps">'
        + _inline_drawing(914400, 914400, "from-choice") +
        '</mc:Choice>'
        '<mc:Fallback>'
        + _vml_pict_rect(50, 30, "FF0000") +
        '</mc:Fallback>'
        '</mc:AlternateContent>'
        '</w:r></w:p>'
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_ac_choice_wins_over_fallback.docx"),
        body, png,
    )

    # v1_ac_fallback_only_pict: ONLY <mc:Fallback><w:pict>. No Choice.
    # Parser falls back to pict path (line 4043-4048). Result is a
    # VML shape (in Paragraph.shapes, NOT an image).
    body = (
        '<w:p><w:r>'
        '<mc:AlternateContent>'
        '<mc:Fallback>'
        + _vml_pict_rect(60, 40, "00FF00") +
        '</mc:Fallback>'
        '</mc:AlternateContent>'
        '</w:r></w:p>'
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_ac_fallback_only_pict.docx"),
        body, png, include_image_rel=False,
    )

    # ===== parse_ole_object =====

    # v1_ole_with_imagedata_preview: <w:object> with a VML shape
    # <v:shape style="width:120pt;height:60pt"><v:imagedata
    # r:id="rIdImg"/></v:shape>. Parser builds an Image with
    # alt_text="OLE Object" HARDCODED, width=120, height=60.
    body = (
        '<w:p><w:r>'
        '<w:object>'
        '<v:shape style="width:120pt;height:60pt">'
        '<v:imagedata r:id="rIdImg"/>'
        '</v:shape>'
        '<o:OLEObject Type="Embed" ProgID="Equation.3" '
        'ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1"/>'
        '</w:object>'
        '</w:r></w:p>'
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_ole_with_imagedata_preview.docx"),
        body, png,
    )

    # v1_ole_no_imagedata: <w:object> with VML shape but NO
    # <v:imagedata>. Parser returns DrawingResult.image = None
    # (line 4011) because rel_id is None. Result: paragraph has
    # NO inline Block::Image.
    body = (
        '<w:p><w:r>'
        '<w:object>'
        '<v:shape style="width:80pt;height:40pt"></v:shape>'
        '<o:OLEObject Type="Embed" ProgID="Equation.3" '
        'ShapeID="_x0000_i1026" DrawAspect="Content" ObjectID="_2"/>'
        '</w:object>'
        '</w:r></w:p>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_ole_no_imagedata.docx"),
               body, png, include_image_rel=False)

    print("Done.")


if __name__ == "__main__":
    main()
