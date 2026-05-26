"""Author minimal `word/theme/theme1.xml` + `word/document.xml`
fixtures for S318 — theme parser end-to-end coverage.

`parse_theme` at crates/oxidocs-core/src/parser/theme.rs:108
extracts:
  - clrScheme entries (dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink)
  - srgbClr val attribute → hex stored on `colors` HashMap
  - sysClr lastClr attribute (NOT val) → hex stored
  - majorFont / minorFont / latin (first-wins): theme.major_font /
    minor_font
  - majorFont / minorFont / ea (first-wins, empty ignored):
    theme.major_font_ea / minor_font_ea
  - <a:font script="Jpan" typeface="..."/> → ea fallback when ea
    not already set
  - End-of-parse fallback: if EA fonts still None, materialize as
    "Meiryo" (line 251-256)

`ThemeColors::resolve` maps Word's themeColor attribute aliases to
internal scheme names:
  dark1|text1 → dk1
  light1|background1 → lt1
  dark2|text2 → dk2
  light2|background2 → lt2
  accent1..6 → passthrough
  hyperlink → hlink
  followedHyperlink → folHlink
  other → passthrough

To make the theme's effect observable from parse_docx, the document
references themeColor / eastAsiaTheme on runs. Then `Run.style.color`
and `Run.style.font_family_east_asia` carry the resolved values
into the IR.

Outputs to ``tools/fixtures/theme_samples/``.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "theme_samples")

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
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

DOC_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
"""

SECT_PR = (
    "<w:sectPr>"
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    "</w:sectPr>"
)


def make_theme_xml(
    dk1_clr: str = '<a:srgbClr val="000000"/>',
    lt1_clr: str = '<a:srgbClr val="FFFFFF"/>',
    dk2_clr: str = '<a:srgbClr val="44546A"/>',
    lt2_clr: str = '<a:srgbClr val="E7E6E6"/>',
    accent1_clr: str = '<a:srgbClr val="4472C4"/>',
    hlink_clr: str = '<a:srgbClr val="0563C1"/>',
    fol_hlink_clr: str = '<a:srgbClr val="954F72"/>',
    major_latin: str = "Calibri Light",
    minor_latin: str = "Calibri",
    major_ea: str = "",
    minor_ea: str = "",
) -> str:
    major_ea_xml = f'<a:ea typeface="{escape(major_ea)}"/>' if major_ea else ""
    minor_ea_xml = f'<a:ea typeface="{escape(minor_ea)}"/>' if minor_ea else ""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="custom">
<a:themeElements>
<a:clrScheme name="custom">
<a:dk1>{dk1_clr}</a:dk1>
<a:lt1>{lt1_clr}</a:lt1>
<a:dk2>{dk2_clr}</a:dk2>
<a:lt2>{lt2_clr}</a:lt2>
<a:accent1>{accent1_clr}</a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink>{hlink_clr}</a:hlink>
<a:folHlink>{fol_hlink_clr}</a:folHlink>
</a:clrScheme>
<a:fontScheme name="custom">
<a:majorFont>
<a:latin typeface="{escape(major_latin)}"/>
{major_ea_xml}
</a:majorFont>
<a:minorFont>
<a:latin typeface="{escape(minor_latin)}"/>
{minor_ea_xml}
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="custom"/>
</a:themeElements>
</a:theme>"""


def _para(rpr_xml: str, text: str) -> str:
    return (
        f'<w:p><w:r><w:rPr>{rpr_xml}</w:rPr>'
        f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
    )


def write_docx(path: str, body_xml: str, theme_xml: str) -> None:
    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
        z.writestr("word/theme/theme1.xml", theme_xml)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_theme_resolve_accent1: themeColor="accent1" + val="FF0000".
    # Theme accent1=00FF00 (custom). Parser at parser/ooxml.rs:4379-4400
    # checks themeColor FIRST; when theme.resolve() returns Some,
    # the resolved hex WINS over the val attribute. So run.color
    # should be "00FF00" (theme), NOT "FF0000" (val fallback).
    theme = make_theme_xml(accent1_clr='<a:srgbClr val="00FF00"/>')
    body = _para(
        '<w:color w:val="FF0000" w:themeColor="accent1"/>',
        "accent1-via-theme",
    )
    write_docx(os.path.join(OUT_DIR, "v1_theme_resolve_accent1.docx"),
               body, theme)

    # v1_theme_resolve_alias_text1: themeColor="text1" → resolve maps
    # to "dk1" (parser/ooxml.rs:25, themes.rs ThemeColors::resolve).
    # Theme dk1=123456. Run.color should be "123456".
    theme = make_theme_xml(dk1_clr='<a:srgbClr val="123456"/>')
    body = _para('<w:color w:themeColor="text1"/>', "text1-resolves-to-dk1")
    write_docx(os.path.join(OUT_DIR, "v1_theme_resolve_alias_text1.docx"),
               body, theme)

    # v1_theme_resolve_hyperlink_alias: themeColor="hyperlink" →
    # resolve maps to "hlink". Theme hlink=0563C1. Distinct alias
    # mapping from text1/dark1 — covers the "hyperlink"/"hlink"
    # match arm at theme.rs:35.
    theme = make_theme_xml(hlink_clr='<a:srgbClr val="0563C1"/>')
    body = _para(
        '<w:color w:themeColor="hyperlink"/>',
        "hyperlink-resolves-to-hlink",
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_theme_resolve_hyperlink_alias.docx"),
        body, theme,
    )

    # v1_theme_sysclr_lastclr: theme uses <a:sysClr lastClr="333333"/>
    # for dk1 (NOT <a:srgbClr val="..."/>). Parser at theme.rs:153-163
    # reads `lastClr` attribute (NOT `val`). A regression that mixed
    # up `val` vs `lastClr` would silently produce empty hex for
    # sysClr-encoded scheme entries.
    theme = make_theme_xml(
        dk1_clr='<a:sysClr val="windowText" lastClr="333333"/>',
    )
    body = _para('<w:color w:themeColor="dark1"/>', "sysClr-lastClr-attr")
    write_docx(os.path.join(OUT_DIR, "v1_theme_sysclr_lastclr.docx"),
               body, theme)

    # v1_theme_font_minor_ea_resolve: theme.minorFont.ea="ＭＳ Ｐ明朝".
    # Run has <w:rFonts w:eastAsiaTheme="minorEastAsia"/>. Parser at
    # parser/ooxml.rs:4269-4276 + styles.rs:15-32 resolves
    # "minorEastAsia" → theme.minor_font_ea → "ＭＳ Ｐ明朝".
    # Without an explicit ea typeface, parse_theme's end-of-parse
    # fallback (theme.rs:251-256) would substitute "Meiryo"; the
    # explicit ea here PREVENTS that fallback.
    theme = make_theme_xml(minor_ea="ＭＳ Ｐ明朝")
    body = _para(
        '<w:rFonts w:eastAsiaTheme="minorEastAsia"/>',
        "ea-theme-resolves",
    )
    write_docx(
        os.path.join(OUT_DIR, "v1_theme_font_minor_ea_resolve.docx"),
        body, theme,
    )

    print("Done.")


if __name__ == "__main__":
    main()
