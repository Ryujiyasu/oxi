"""Author minimal word/fontTable.xml repro fixtures.

Document.styles.font_table end-to-end coverage (S303): parser at
parser/ooxml.rs:475 (parse_font_table) reads word/fontTable.xml and
populates Document.styles.font_table: HashMap<String, FontInfo>, where
each FontInfo carries optional panose1, charset, family, pitch — all
verbatim string values keyed by the parent <w:font w:name="...">.

Unit tests cover XML parsing; these fixtures verify the full
parse_docx → Document.styles.font_table roundtrip including:
  - HashMap keying by w:name
  - Per-field Option<String> when child element absent
  - Verbatim preservation of panose1 hex (20 chars, no normalization)
  - Multiple <w:font> entries each becoming its own HashMap entry

Outputs to ``tools/fixtures/font_info_samples/`` (committed,
S272 no-COM direct-write pattern).

Fixtures (5):
  v1_basic.docx             — 1 font with all 4 fields populated
  v1_partial.docx           — 1 font with only panose1 (others None)
  v1_multiple_fonts.docx    — 3 fonts keyed independently
  v1_panose_verbatim.docx   — 20-char PANOSE hex preserved exactly
  v1_no_fonttable.docx      — fontTable.xml absent → empty font_table
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "font_info_samples")

CONTENT_TYPES_WITH_FONTS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
</Types>"""

CONTENT_TYPES_NO_FONTS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t xml:space="preserve">Font info fixture body.</w:t></w:r></w:p>
"""

SECT_PR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '</w:sectPr>'
)

FONT_TABLE_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
"""

FONT_TABLE_TAIL = "</w:fonts>"


def _font(name: str, panose1: str | None = None, charset: str | None = None,
          family: str | None = None, pitch: str | None = None) -> str:
    """Build a <w:font w:name=...><w:panose1.../>...<w:pitch.../></w:font> block.
    Each field is emitted only when its value is provided."""
    children = ""
    if panose1 is not None:
        children += f'<w:panose1 w:val="{panose1}"/>'
    if charset is not None:
        children += f'<w:charset w:val="{charset}"/>'
    if family is not None:
        children += f'<w:family w:val="{family}"/>'
    if pitch is not None:
        children += f'<w:pitch w:val="{pitch}"/>'
    return f'<w:font w:name="{escape(name)}">{children}</w:font>'


def write_docx(path: str, font_table_xml: str | None) -> None:
    body_xml = DOC_HEAD + SECT_PR + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        if font_table_xml is not None:
            z.writestr("[Content_Types].xml", CONTENT_TYPES_WITH_FONTS)
        else:
            z.writestr("[Content_Types].xml", CONTENT_TYPES_NO_FONTS)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", body_xml)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
        if font_table_xml is not None:
            z.writestr("word/fontTable.xml", font_table_xml)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_basic: 1 font with all 4 fields
    ft = FONT_TABLE_HEAD + _font(
        "Calibri",
        panose1="020F0502020204030204",
        charset="00",
        family="swiss",
        pitch="variable",
    ) + FONT_TABLE_TAIL
    write_docx(os.path.join(OUT_DIR, "v1_basic.docx"), ft)

    # v1_partial: 1 font with only panose1 (charset/family/pitch absent)
    ft = FONT_TABLE_HEAD + _font(
        "ＭＳ 明朝",
        panose1="02020609040205080304",
    ) + FONT_TABLE_TAIL
    write_docx(os.path.join(OUT_DIR, "v1_partial.docx"), ft)

    # v1_multiple_fonts: 3 fonts keyed independently
    ft = FONT_TABLE_HEAD + _font(
        "Times New Roman", panose1="02020603050405020304",
        charset="00", family="roman", pitch="variable",
    ) + _font(
        "Courier New", panose1="02070309020205020404",
        charset="00", family="modern", pitch="fixed",
    ) + _font(
        "Wingdings", panose1="05000000000000000000",
        charset="02", family="auto", pitch="default",
    ) + FONT_TABLE_TAIL
    write_docx(os.path.join(OUT_DIR, "v1_multiple_fonts.docx"), ft)

    # v1_panose_verbatim: 20-char PANOSE hex preserved exactly
    # Use a non-standard but valid-looking hex string so we can verify
    # the parser doesn't normalize/case-fold/strip.
    ft = FONT_TABLE_HEAD + _font(
        "PanoseTest",
        panose1="abcdef0123456789ABCD",
    ) + FONT_TABLE_TAIL
    write_docx(os.path.join(OUT_DIR, "v1_panose_verbatim.docx"), ft)

    # v1_no_fonttable: fontTable.xml absent → font_table HashMap empty
    write_docx(os.path.join(OUT_DIR, "v1_no_fonttable.docx"), None)

    print("Done.")


if __name__ == "__main__":
    main()
