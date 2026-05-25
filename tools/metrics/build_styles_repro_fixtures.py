"""Author minimal styles.xml repro fixtures for end-to-end stylesheet coverage (S298).

`StyleSheet` end-to-end coverage:
 - parser at parser/styles.rs:36 `parse_styles` walks `<w:docDefaults>` and
   each `<w:style>` block, then `resolve_style_inheritance` flattens
   `<w:basedOn>` chains by merging parent ParagraphStyle / RunStyle into
   each child.
 - Unit tests in `crates/oxidocs-core/src/parser/styles.rs` cover XML-level
   parsing of individual blocks; these fixtures verify the full
   `parse_docx` → `Document.styles.{doc_default_*,default_paragraph_style_id,styles}`
   roundtrip across the four common stylesheet-shape combinations.

Outputs to ``tools/fixtures/styles_samples/`` directly (committed,
S272 no-COM direct-write pattern).

Fixtures (4):
  v1_doc_defaults.docx          docDefaults only:
                                rPrDefault Calibri sz=24 (=12pt) +
                                pPrDefault jc=center, spacing line=276 (auto)
  v1_default_para_style.docx    one paragraph style id="Normal"
                                with type="paragraph" default="1"
  v1_basedon_chain.docx         Heading1 basedOn Normal:
                                Normal has rPr sz=22 (=11pt) only,
                                Heading1 has rPr <w:b/> only
                                → resolved Heading1 has BOTH bold and 11pt
  v1_para_style_alignment.docx  style id="CenteredIndent" with
                                pPr jc=center + ind left=720 (=36pt)
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures", "styles_samples")

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

# Standard A4 sectPr with 72pt margins, shared across all fixtures.
SECT_PR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '</w:sectPr>'
)

DOC_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
"""


def _paragraph(text: str, style_id: str | None = None) -> str:
    pstyle = f'<w:pPr><w:pStyle w:val="{style_id}"/></w:pPr>' if style_id else ''
    return (
        f'<w:p>{pstyle}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
    )


def write_docx(path: str, body_xml: str, styles_xml: str) -> None:
    full = DOC_HEAD + body_xml + SECT_PR + "\n</w:body>\n</w:document>"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", styles_xml)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


# v1_doc_defaults: docDefaults only.
#   rPrDefault: Calibri ascii/hAnsi + sz=24 (=12pt) + color=2E74B5
#   pPrDefault: jc=center + spacing line=276 (auto → 1.15 multiplier)
STYLES_V1_DOC_DEFAULTS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="MS Mincho"/>
<w:sz w:val="24"/>
<w:color w:val="2E74B5"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr>
<w:jc w:val="center"/>
<w:spacing w:line="276" w:lineRule="auto"/>
</w:pPr></w:pPrDefault>
</w:docDefaults>
</w:styles>"""

# v1_default_para_style: single paragraph style id="Normal" with default="1".
# No docDefaults at all. Tests that default_paragraph_style_id capture is
# driven by the w:default="1" attribute, not by the style id name.
STYLES_V1_DEFAULT_PARA_STYLE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal" w:default="1">
<w:name w:val="Normal"/>
<w:qFormat/>
</w:style>
</w:styles>"""

# v1_basedon_chain: Heading1 basedOn Normal.
#   Normal: rPr.sz=22 (=11pt)
#   Heading1: rPr.<w:b/>
# After resolve_style_inheritance the merged Heading1 default_run_style
# should have BOTH bold=true (from itself) AND font_size=11.0 (from Normal).
STYLES_V1_BASEDON_CHAIN = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal" w:default="1">
<w:name w:val="Normal"/>
<w:rPr><w:sz w:val="22"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Heading1">
<w:name w:val="heading 1"/>
<w:basedOn w:val="Normal"/>
<w:rPr><w:b/></w:rPr>
</w:style>
</w:styles>"""

# v1_para_style_alignment: style id="CenteredIndent" with pPr only.
#   pPr.jc=center → StyleDefinition.alignment = Some(Alignment::Center)
#   pPr.ind w:left=720 (=36pt) → ParagraphStyle.indent_left = Some(36.0)
STYLES_V1_PARA_STYLE_ALIGNMENT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="CenteredIndent">
<w:name w:val="Centered Indent"/>
<w:pPr>
<w:jc w:val="center"/>
<w:ind w:left="720"/>
</w:pPr>
</w:style>
</w:styles>"""


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    write_docx(
        os.path.join(OUT_DIR, "v1_doc_defaults.docx"),
        _paragraph("docDefaults body."),
        STYLES_V1_DOC_DEFAULTS,
    )

    write_docx(
        os.path.join(OUT_DIR, "v1_default_para_style.docx"),
        _paragraph("Default-paragraph-style body."),
        STYLES_V1_DEFAULT_PARA_STYLE,
    )

    write_docx(
        os.path.join(OUT_DIR, "v1_basedon_chain.docx"),
        _paragraph("Heading body.", style_id="Heading1"),
        STYLES_V1_BASEDON_CHAIN,
    )

    write_docx(
        os.path.join(OUT_DIR, "v1_para_style_alignment.docx"),
        _paragraph("Centered indent body.", style_id="CenteredIndent"),
        STYLES_V1_PARA_STYLE_ALIGNMENT,
    )

    print("Done.")


if __name__ == "__main__":
    main()
