"""Author minimal `<w:pict>` (VML legacy shape) repro fixtures for S317.

`parse_vml_pict` at parser/ooxml.rs:3712 handles the LEGACY VML
shape format (<v:rect>, <v:shape>, <v:oval>, <v:roundrect>, <v:line>)
that pre-DrawingML docs use. This is distinct from `parse_drawing`
(parser/ooxml.rs:2958) which handles the modern <w:drawing> /
<wp:inline|anchor> path.

Surface covered:
  - VML shape elements → Shape.shape_type string. Five distinct
    map arms at parser/ooxml.rs:3758-3773 plus the special-case
    t185 → "bracketPair" (CLAUDE.md S70 double-bracket 〔〕).
  - CSS-like `style` attribute on shape elements parsed at line
    3779-3795. Five-way unit conversion in parse_css_length:
    "pt"→val, "in"→val*72, "cm"→val*28.3465, "mm"→val*2.83465,
    "px"→val*0.75 (96dpi→72pt), no-suffix→val raw.
  - Boolean polarity: `filled="f"` or `filled="false"` → no_fill=true.
    Same idiom for `stroked`. Line 3797, 3801.
  - Color leading `#` strip: `fillcolor="#FF0000"` → fill="FF0000"
    via `.trim_start_matches('#')`. Pinning catches a regression
    that lost the strip and surfaces a "#FF0000" downstream.
  - Absolute position routing: `style="position:absolute;...;
    margin-left:X;margin-top:Y"` → Shape.position = Some with
    h_relative=v_relative="text" (HARDCODED — VML uses text anchor
    only). Line 3876-3887.
  - stroke_width fallback: when stroked is enabled and no
    strokeweight specified, Shape.stroke_width defaults to
    Some(0.75) (line 3895). NOT None.

Outputs to ``tools/fixtures/vml_pict_samples/``.
"""
import os
import zipfile

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "vml_pict_samples")

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
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

# Namespaces needed for VML: v (urn:schemas-microsoft-com:vml).
DOC_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
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


def _pict_paragraph(vml_shape_xml: str) -> str:
    return f'<w:p><w:r><w:pict>{vml_shape_xml}</w:pict></w:r></w:p>'


def write_docx(path: str, body_xml: str) -> None:
    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # IMPORTANT: parse_vml_pict at parser/ooxml.rs:3758 matches shape
    # elements on `Event::Start` ONLY, NOT `Event::Empty`. Self-closing
    # `<v:rect ... />` triggers Empty and the shape is silently
    # dropped. All fixtures must use explicit start+end pairs
    # (`<v:rect ...></v:rect>`) to exercise the shape-recognition
    # path. This is a latent parser asymmetry pinned by these
    # fixtures — a future refactor that "normalizes" self-closing
    # to expanded form would change observed behavior.

    # v1_vml_rect_basic: <v:rect> with fill/stroke colors (leading #
    # stripped) + strokeweight + v-text-anchor.
    shape = (
        '<v:rect style="width:100pt;height:50pt;v-text-anchor:middle" '
        'fillcolor="#FF0000" strokecolor="#0000FF" strokeweight="2pt">'
        '</v:rect>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_vml_rect_basic.docx"),
               _pict_paragraph(shape))

    # v1_vml_bracket_pair_t185: <v:shape type="...t185"> → "bracketPair".
    # CLAUDE.md S70 2026-04-13: VML preset t185 is the double-bracket
    # 〔〕 shape. The t185 substring is the ONLY path in the VML parser
    # that distinguishes a specific preset; everything else falls
    # through to "rect".
    shape = (
        '<v:shape type="#_x0000_t185" '
        'style="width:50pt;height:30pt"></v:shape>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_vml_bracket_pair_t185.docx"),
               _pict_paragraph(shape))

    # v1_vml_oval_no_fill_no_stroke: <v:oval> with filled="f" +
    # stroked="f" → both polarity-flipped. Shape.fill=None AND
    # stroke_color=None AND stroke_width=None (stroked=false
    # SUPPRESSES the 0.75 default).
    shape = (
        '<v:oval style="width:120pt;height:60pt" '
        'fillcolor="#AAAAAA" strokecolor="#000000" '
        'filled="f" stroked="f"></v:oval>'
    )
    write_docx(os.path.join(OUT_DIR, "v1_vml_oval_no_fill_no_stroke.docx"),
               _pict_paragraph(shape))

    # v1_vml_roundrect_absolute_position: <v:roundrect> with
    # position:absolute + margin-left + margin-top → Shape.position
    # = Some(FloatingPosition {h_relative:"text", v_relative:"text"}).
    # HARDCODED text anchor — DIFFERENT from DrawingML positionH/V.
    shape = (
        '<v:roundrect '
        'style="position:absolute;width:80pt;height:40pt;'
        'margin-left:50pt;margin-top:30pt" '
        'fillcolor="#00FF00"></v:roundrect>'
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_vml_roundrect_absolute_position.docx"),
        _pict_paragraph(shape),
    )

    # v1_vml_css_units: mix of CSS units in one shape style.
    # width:1in → 72pt (val * 72.0). height:36pt → 36pt (raw).
    # No filled="f"/stroked="f" → stroke_width defaults to 0.75.
    shape = '<v:rect style="width:1in;height:36pt"></v:rect>'
    write_docx(os.path.join(OUT_DIR, "v1_vml_css_units.docx"),
               _pict_paragraph(shape))

    print("Done.")


if __name__ == "__main__":
    main()
