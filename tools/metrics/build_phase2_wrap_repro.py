"""Author minimal repro fixtures isolating the Phase 2 wrap-count bug
class identified in S299/S300.

Bug class: in narrow table cells with exact line height and CJK text,
Word and Oxi disagree about wrap break positions. Direction can be
EITHER way depending on content (29dc6e: Oxi packs more, b35123:
Oxi packs less).

This builder authors three minimal repros, each isolating ONE plausible
root cause, so COM measurement on each can identify which is at play:

  v1_cjk_only          plain CJK text, no parens / no half-width digits
                       → tests the base CJK char-width
  v2_cjk_paren_digit   mixes CJK with `（` `）` and `77` half-width digits
                       (matches 29dc6e i=265 content profile)
                       → tests punctuation / half-width width
  v3_cjk_latin_punc    mixes CJK with `:` `,` `;` Latin punctuation and
                       short Latin words (matches e3c545 page 4 profile)
                       → tests Latin/CJK boundary handling

All three use the SAME cell:
  tcW = 7541 dxa = 377.05pt (matches 29dc6e i=265 cell exactly)
  gridSpan = 1 (single-column repro — eliminates gridSpan confound)
  exact 11pt line height (w:line=220, w:lineRule=exact)
  hanging indent 217 (10.85pt)
  left indent 543 (27.15pt), right 156 (7.8pt)

The text in each variant is calibrated to wrap to ~3 lines in Word.
If Oxi wraps to a different number of lines on any variant, that
variant's char class is the culprit.

Outputs to tools/fixtures/phase2_wrap_samples/ for reuse across sessions.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(
    os.path.dirname(__file__), "..", "fixtures", "phase2_wrap_samples"
)

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

# Use MS Mincho 11pt for Japanese (sz=22 in halfpoints).
# docDefaults match Word's typical defaults for tokumei-class docs.
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝"/>
<w:sz w:val="22"/><w:szCs w:val="22"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
</w:styles>"""

SECT_PR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '</w:sectPr>'
)

# Replicate the 29dc6e cell exactly:
#   tcW=7541tw = 377.05pt
#   pPr: <w:spacing w:before="120" w:line="220" w:lineRule="exact"/>
#        <w:ind w:left="543" w:right="156" w:hanging="217"/>
# Single-column gridSpan=1 (eliminates the 7-column-span confound).
def cell_table_xml(paragraph_text: str) -> str:
    return (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="7541" w:type="dxa"/>'
        '<w:tblInd w:w="12" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
        '<w:tblLayout w:type="fixed"/>'
        '</w:tblPr>'
        '<w:tblGrid>'
        '<w:gridCol w:w="7541"/>'
        '</w:tblGrid>'
        '<w:tr>'
        '<w:tc>'
        '<w:tcPr>'
        '<w:tcW w:w="7541" w:type="dxa"/>'
        '<w:tcBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tcBorders>'
        '</w:tcPr>'
        '<w:p>'
        '<w:pPr>'
        '<w:spacing w:before="120" w:line="220" w:lineRule="exact"/>'
        '<w:ind w:left="543" w:right="156" w:hanging="217"/>'
        '<w:rPr><w:spacing w:val="0"/></w:rPr>'
        '</w:pPr>'
        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:spacing w:val="0"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(paragraph_text)}</w:t>'
        '</w:r>'
        '</w:p>'
        '</w:tc>'
        '</w:tr>'
        '</w:tbl>'
    )


def write_docx(path: str, paragraph_text: str) -> None:
    body_xml = cell_table_xml(paragraph_text)
    full = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '<w:body>\n'
        + body_xml +
        SECT_PR +
        '\n</w:body>\n</w:document>'
    )
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


# Variant 1: plain CJK ~80 chars, no parens, no half-width digits.
# Tests pure CJK char-width measurement under exact 11pt line height.
V1_TEXT = (
    "○　"
    "暴力団員等がその事業活動を支配する者又は暴力団員等をその"
    "業務に従事させ若しくは当該業務の補助者として使用するおそれのある者"
)

# Variant 2: mix of CJK + fullwidth digit ３ + halfwidth digit "77" +
# fullwidth parens （ ） + 、 。 punctuation.
# Matches 29dc6e i=265 content profile EXACTLY.
V2_TEXT = (
    "○　"
    "暴力団員による不当な行為の防止等に関する法律"
    "（平成３年法律第77号）"
    "第２条第６号に規定する暴力団員又は暴力団員でなくなった日"
    "から５年を経過しない者（以下「暴力団員等」という。）"
)

# Variant 3: CJK + Latin words + ASCII colon/semicolon/comma.
# Matches e3c545 page 4 content profile (mixed Latin/CJK in code-like
# table cell).
V3_TEXT = (
    "□　"
    "rdf:type void:Dataset ; dcterms:title \"xxx-city stats\" ; "
    "dcterms:creator <http://example.com> ;"
    "dcterms:created \"2016-03-01\" ; cc:license <http://creativecommons.org/>"
)


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")
    write_docx(os.path.join(OUT_DIR, "v1_cjk_only.docx"),         V1_TEXT)
    write_docx(os.path.join(OUT_DIR, "v2_cjk_paren_digit.docx"),  V2_TEXT)
    write_docx(os.path.join(OUT_DIR, "v3_cjk_latin_punc.docx"),   V3_TEXT)
    print("Done.")
    print()
    print(f"Variant lengths (chars):")
    print(f"  V1 (pure CJK):           {len(V1_TEXT)}")
    print(f"  V2 (CJK+paren+digit):    {len(V2_TEXT)}")
    print(f"  V3 (CJK+Latin+punc):     {len(V3_TEXT)}")


if __name__ == "__main__":
    main()
