"""Author minimal inline-image repro fixtures with embedded PNG bytes.

Image (Block::Image) end-to-end coverage (S292): parser at
parser/ooxml.rs:359 reads image relationships from
word/_rels/document.xml.rels and loads the binary file from
word/media/imageN.{png,jpg,...}. parser/ooxml.rs:2993 extracts
`docPr/@descr` as `Image.alt_text`. parser/ooxml.rs:3407 reads
`wp:extent/@cx,@cy` (EMU) as width/height. Unit tests cover XML
parsing; these fixtures verify the full parse_docx → Document walk →
Block::Image roundtrip including binary blob.

Outputs to ``tools/fixtures/image_samples/`` directly.

Fixtures (3):
  v1_simple.docx       — inline 1×1 PNG, alt_text="Test alt text"
  v1_no_alt.docx       — image without descr attribute (alt_text None)
  v1_custom_size.docx  — image with cx=1828800 cy=914400 EMU (2in×1in)
"""
import os
import struct
import zlib
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "image_samples")


def minimal_png(w: int = 1, h: int = 1, rgba=(255, 255, 255, 255)) -> bytes:
    """Construct a valid w×h RGBA PNG with all pixels = rgba."""
    r, g, b, a = rgba
    sig = b'\x89PNG\r\n\x1a\n'
    # IHDR
    ihdr_data = struct.pack('>IIBBBBB', w, h, 8, 6, 0, 0, 0)
    ihdr_chunk = b'IHDR' + ihdr_data
    ihdr = struct.pack('>I', len(ihdr_data)) + ihdr_chunk + struct.pack('>I', zlib.crc32(ihdr_chunk))
    # IDAT: per scanline, prepend filter byte (0 = None) then 4 bytes per pixel
    raw = b''
    for _ in range(h):
        raw += bytes([0]) + bytes([r, g, b, a]) * w
    compressed = zlib.compress(raw)
    idat_chunk = b'IDAT' + compressed
    idat = struct.pack('>I', len(compressed)) + idat_chunk + struct.pack('>I', zlib.crc32(idat_chunk))
    # IEND
    iend_chunk = b'IEND'
    iend = struct.pack('>I', 0) + iend_chunk + struct.pack('>I', zlib.crc32(iend_chunk))
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
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '</w:sectPr>'
)


def _inline_image_run(cx: int, cy: int, descr: str | None = None,
                     name: str = "Picture 1", img_rel_id: str = "rIdImg") -> str:
    descr_attr = f' descr="{escape(descr)}"' if descr is not None else ''
    return f"""<w:r><w:drawing>
<wp:inline distT="0" distB="0" distL="0" distR="0">
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:docPr id="1" name="{escape(name)}"{descr_attr}/>
<wp:cNvGraphicFramePr/>
<a:graphic>
<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic>
<pic:nvPicPr>
<pic:cNvPr id="1" name="{escape(name)}"/>
<pic:cNvPicPr/>
</pic:nvPicPr>
<pic:blipFill>
<a:blip r:embed="{img_rel_id}"/>
<a:stretch><a:fillRect/></a:stretch>
</pic:blipFill>
<pic:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</pic:spPr>
</pic:pic>
</a:graphicData>
</a:graphic>
</wp:inline>
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
        z.writestr("word/settings.xml", SETTINGS_Xml := SETTINGS_XML)  # use both
        z.writestr("word/media/image1.png", png_bytes)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")
    png_1x1 = minimal_png(1, 1)
    print(f"  minimal PNG bytes: {len(png_1x1)}")

    # v1_simple: 1×1 PNG with alt_text, cx=cy=914400 EMU (1 inch each)
    body = _paragraph(_inline_image_run(914400, 914400, descr="Test alt text"))
    write_docx(os.path.join(OUT_DIR, "v1_simple.docx"), body, png_1x1)

    # v1_no_alt: image without descr attribute → alt_text should be None
    body = _paragraph(_inline_image_run(914400, 914400, descr=None))
    write_docx(os.path.join(OUT_DIR, "v1_no_alt.docx"), body, png_1x1)

    # v1_custom_size: image with cx=1828800 cy=914400 EMU = 2×1 inch
    # EMU/914400 = inch, EMU/12700 = pt. So 1828800/12700 = 144pt width,
    # 914400/12700 = 72pt height
    body = _paragraph(_inline_image_run(1828800, 914400, descr="Custom-sized image"))
    write_docx(os.path.join(OUT_DIR, "v1_custom_size.docx"), body, png_1x1)

    print("Done.")


if __name__ == "__main__":
    main()
