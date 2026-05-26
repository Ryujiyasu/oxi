"""S323 minimal repro: Yu Mincho 11pt via theme vs explicit rFonts.

Hypothesis from S322: Word applies 83/64 multiplier to Yu Mincho when
the run's rFonts is explicit (`<w:rFonts w:eastAsia="游明朝"/>`) but
NOT when the Yu Mincho came via theme.minorEastAsia resolution
(`<w:rFonts eastAsiaTheme="minorEastAsia"/>`).

Test: build TWO minimal docx that differ ONLY in how the run's
Yu Mincho font is specified:

  v1_yu_mincho_theme.docx     — eastAsiaTheme="minorEastAsia"
                                 + theme.minorFont.ea="游明朝"
  v1_yu_mincho_explicit.docx  — eastAsia="游明朝" (no theme)

Both should produce IDENTICAL run.style.font_family_east_asia="游明朝"
in Oxi's IR. The question is what Word renders.

Outputs to ``tools/fixtures/yu_mincho_theme_vs_explicit/``.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "yu_mincho_theme_vs_explicit")

# Copy from a working real docx (d1e8ac8) so Word accepts our fixtures.
SRC_DOCX = os.path.join(
    os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "d1e8ac8fd1cc_kyodokenkyuyoushiki06.docx",
)


def _read_src_part(name: str) -> bytes:
    with zipfile.ZipFile(SRC_DOCX) as z:
        return z.read(name)

CONTENT_TYPES_THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

DOC_RELS_THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

# docDefault uses ＭＳ 明朝 — same as d1e8ac8 — so theme's Yu Mincho
# is reached only via explicit eastAsiaTheme on the run.
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
</w:styles>"""

# Same theme as d1e8ac8: minorFont.latin=游明朝 (Yu Mincho),
# Jpan script also Yu Mincho.
THEME_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="custom">
<a:themeElements>
<a:clrScheme name="custom">
<a:dk1><a:srgbClr val="000000"/></a:dk1>
<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2>
<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="custom">
<a:majorFont>
<a:latin typeface="游ゴシック Light"/>
<a:ea typeface=""/>
<a:font script="Jpan" typeface="游ゴシック Light"/>
</a:majorFont>
<a:minorFont>
<a:latin typeface="游明朝"/>
<a:ea typeface=""/>
<a:font script="Jpan" typeface="游明朝"/>
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="custom"/>
</a:themeElements>
</a:theme>"""

DOC_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">
<w:body>
"""

SECT_PR = (
    "<w:sectPr>"
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    "</w:sectPr>"
)


def _para(rfonts_xml: str, text: str) -> str:
    # Use 3 paragraphs so the FIRST is a control (no rFonts, baseline)
    # and the SECOND/THIRD are the test cases. Word's Information(6)
    # measures Y per paragraph so we can read the stride directly.
    return (
        f'<w:p><w:r><w:rPr>{rfonts_xml}<w:sz w:val="22"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
    )


def write_docx(path: str, body_xml: str, with_theme: bool) -> None:
    # Use REAL theme/styles/settings from d1e8ac8 (so Yu Mincho theme
    # resolution works correctly) but minimal Content_Types and
    # document.xml.rels referencing only what we actually include.
    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    src_styles = _read_src_part("word/styles.xml")
    src_settings = _read_src_part("word/settings.xml")
    src_theme = _read_src_part("word/theme/theme1.xml")

    minimal_content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '<Default Extension="xml" ContentType="application/xml"/>\n'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
        '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>\n'
        '</Types>'
    )

    minimal_root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n'
        '</Relationships>'
    )

    minimal_doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>\n'
        '</Relationships>'
    )

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", minimal_content_types)
        z.writestr("_rels/.rels", minimal_root_rels)
        z.writestr("word/_rels/document.xml.rels", minimal_doc_rels)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", src_styles)
        z.writestr("word/settings.xml", src_settings)
        z.writestr("word/theme/theme1.xml", src_theme)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_yu_mincho_theme: 3 paragraphs, all 11pt MS Mincho except the
    # MIDDLE one uses Yu Mincho via theme.minorEastAsia. Three paras
    # so we can measure stride from p1→p2 (MS Mincho), p2→p3 (Yu
    # via theme), p3→pageEnd (MS Mincho).
    body = (
        _para('<w:rFonts w:hint="eastAsia"/>', "あいう")
        + _para('<w:rFonts w:asciiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia" w:hint="eastAsia"/>',
                "テーマ")
        + _para('<w:rFonts w:hint="eastAsia"/>', "末尾")
    )
    write_docx(os.path.join(OUT_DIR, "v1_yu_mincho_theme.docx"),
               body, with_theme=True)

    # v1_yu_mincho_explicit: middle paragraph uses Yu Mincho via
    # EXPLICIT rFonts eastAsia="游明朝" (no theme indirection).
    body = (
        _para('<w:rFonts w:hint="eastAsia"/>', "あいう")
        + _para('<w:rFonts w:ascii="游明朝" w:eastAsia="游明朝" w:hAnsi="游明朝" w:hint="eastAsia"/>',
                "テーマ")
        + _para('<w:rFonts w:hint="eastAsia"/>', "末尾")
    )
    write_docx(os.path.join(OUT_DIR, "v1_yu_mincho_explicit.docx"),
               body, with_theme=True)

    # v1_ms_mincho_only: all 3 paragraphs MS Mincho (control — no
    # Yu involvement, so all 3 strides should match Word's baseline
    # behavior).
    body = (
        _para('<w:rFonts w:hint="eastAsia"/>', "あいう")
        + _para('<w:rFonts w:hint="eastAsia"/>', "あいう")
        + _para('<w:rFonts w:hint="eastAsia"/>', "末尾")
    )
    write_docx(os.path.join(OUT_DIR, "v1_ms_mincho_only.docx"),
               body, with_theme=True)

    print("Done.")
    print()
    print("Next: open each fixture in Word, measure paragraph Y via")
    print("  doc.Paragraphs(N).Range.Information(6)")
    print("and compute stride. Or render via Oxi gdi renderer and")
    print("compare to existing baseline.")


if __name__ == "__main__":
    main()
