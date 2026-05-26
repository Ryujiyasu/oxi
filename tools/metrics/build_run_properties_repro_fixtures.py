"""Author minimal `<w:r><w:rPr>...</w:rPr>` repro fixtures for S308
run-level RunStyle coverage.

`styles_integration.rs` (S298) covers `<w:rPrDefault>` and style-sheet
`<w:rPr>` end-to-end; `comments_fixtures.rs::fixture_09` exercises
`Run.style.bold` via an rPrChange toggle. But no integration test
exercises the inline `<w:r><w:rPr>...</w:rPr>` parser path at
parser/ooxml.rs:4189 (`parse_run_properties`) across the breadth of
fields it handles: half-point sz arithmetic, dstrike → BOTH strike
+ double_strikethrough, vertAlign val=subscript/superscript,
color val="auto" → None (not "auto"), highlight enum value, w:w
text_scale percentage, kern (half-points → pt), position (half-points
→ pt), spacing (twips → pt), rFonts ascii vs eastAsia separation,
has_explicit_east_asia=true when eastAsia is set as an explicit
attribute, webHidden as an alias for vanish, run-level shd fill, and
smallCaps/caps as independent flags.

Outputs to ``tools/fixtures/run_properties_samples/``.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "run_properties_samples")

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

# Minimal styles.xml — no docDefaults so the inline rPr is the SOLE source
# of run formatting. This isolates parse_run_properties from any merge
# behavior that would mask whether a given field came from the inline
# rPr vs a default.
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:styles>"""

DOC_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
"""

DOC_TAIL = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '</w:sectPr>'
    '\n</w:body>\n</w:document>'
)


def _para(runs_xml: str) -> str:
    return f'<w:p>{runs_xml}</w:p>'


def _run(text: str, rpr_xml: str = "") -> str:
    rpr = f'<w:rPr>{rpr_xml}</w:rPr>' if rpr_xml else ''
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def write_docx(path: str, body_xml: str) -> None:
    full = DOC_HEAD + body_xml + DOC_TAIL
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

    # v1_bold_italic_underline: pins three boolean flags + the
    # `<u w:val="none"/>` polarity flip (val=none SUPPRESSES underline
    # even though the element is present — a non-obvious branch at
    # parser/ooxml.rs:4291).
    body = _para(
        _run("plain") +
        _run("bold-italic", '<w:b/><w:i/>') +
        _run("underline-single", '<w:u w:val="single"/>') +
        _run("underline-none", '<w:u w:val="none"/>')
    )
    write_docx(os.path.join(OUT_DIR, "v1_bold_italic_underline.docx"), body)

    # v1_strike_dstrike_vertalign: pins dstrike → BOTH strikethrough
    # AND double_strikethrough (parser/ooxml.rs:4300-4303), and
    # vertAlign val=subscript/superscript → VerticalAlign enum.
    body = _para(
        _run("plain") +
        _run("strike", '<w:strike/>') +
        _run("dstrike", '<w:dstrike/>') +
        _run("super", '<w:vertAlign w:val="superscript"/>') +
        _run("sub", '<w:vertAlign w:val="subscript"/>')
    )
    write_docx(os.path.join(OUT_DIR, "v1_strike_dstrike_vertalign.docx"), body)

    # v1_sz_color_highlight: pins half-point sz arithmetic (val=22 → 11.0,
    # val=23 → 11.5 NOT rounded), color val="auto" → None (suppression,
    # not stored verbatim), color val=hex → Some(hex), highlight stored
    # as the enum-string verbatim.
    body = _para(
        _run("sz22", '<w:sz w:val="22"/>') +
        _run("sz23-half", '<w:sz w:val="23"/>') +
        _run("color-auto", '<w:color w:val="auto"/>') +
        _run("color-red", '<w:color w:val="FF0000"/>') +
        _run("highlight-yellow", '<w:highlight w:val="yellow"/>')
    )
    write_docx(os.path.join(OUT_DIR, "v1_sz_color_highlight.docx"), body)

    # v1_kern_position_spacing: pins twip→pt and half-point→pt conversion
    # for character_spacing (twips/20), kern (half-points/2), position
    # (half-points/2, signed — `val=-6` → -3.0pt lowered), and w:w
    # text_scale (raw percentage, no conversion).
    body = _para(
        _run("kern22", '<w:kern w:val="22"/>') +
        _run("pos-raised", '<w:position w:val="6"/>') +
        _run("pos-lowered", '<w:position w:val="-6"/>') +
        _run("spacing-40tw", '<w:spacing w:val="40"/>') +
        _run("scale80", '<w:w w:val="80"/>')
    )
    write_docx(os.path.join(OUT_DIR, "v1_kern_position_spacing.docx"), body)

    # v1_rfonts_caps_vanish: pins
    #   - rFonts ascii AND eastAsia → font_family AND font_family_east_asia
    #     populated independently, and has_explicit_east_asia=true.
    #   - smallCaps and caps are INDEPENDENT flags (both can be true).
    #   - webHidden is an alias for vanish (parser/ooxml.rs:4445 treats
    #     both as the same flag — NOT a separate field).
    #   - shd fill = run-level character shading (separate from paragraph
    #     shading).
    body = _para(
        _run(
            "arial-cjk",
            '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="ＭＳ Ｐ明朝"/>',
        ) +
        _run("smallcaps-caps", '<w:smallCaps/><w:caps/>') +
        _run("hidden-webhidden", '<w:webHidden/>') +
        _run("hidden-vanish", '<w:vanish/>') +
        _run("shaded", '<w:shd w:fill="FFFF00"/>')
    )
    write_docx(os.path.join(OUT_DIR, "v1_rfonts_caps_vanish.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
