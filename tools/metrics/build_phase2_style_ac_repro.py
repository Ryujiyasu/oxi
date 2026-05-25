"""S300 — Build a minimal repro that DOES reproduce the 29dc6e wrap-count bug.

S299 established (COM-confirmed):
  Word renders 29dc6e i=265 in 3 line slots (line_in_page 31->34 = 3).
  Oxi renders it in 2 lines.
  Wrap-count differs by 1.

S299 also established that the V1 minimal repro (sz=22, no style "ac")
did NOT reproduce — Oxi happily wrapped it to 3 lines, same as Word.

Hypothesis for the reproducer: applying full style "ac" (which the full
29dc6e doc uses) flips Oxi's wrap to 2 lines while Word stays at 3.

Style "ac" pPr properties:
  widowControl=0, wordWrap=0, autoSpaceDE=0, autoSpaceDN=0,
  adjustRightInd=0, spacing line=210 exact, jc=both
Style "ac" rPr properties:
  rFonts ascii="ＭＳ 明朝" hAnsi="ＭＳ 明朝" cs="ＭＳ 明朝"
  spacing val=-1 (char spacing -1tw = -0.05pt per char)
  sz=21 szCs=21 (10.5pt)

Paragraph then overrides:
  spacing line=220 exact (11pt, vs style's 10.5pt)
  ind left=543 right=156 hanging=217
  run rPr spacing val=0 (overrides style's -1)

Fixture v4 reproduces this faithfully. Subsequent v4a/b/c/d will hold
v4 as the bug-repro baseline and toggle each candidate property to
identify the differentiator.
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


def styles_xml(*,
               word_wrap: str = "0",
               auto_space_de: str = "0",
               auto_space_dn: str = "0",
               char_spacing: str = "-1") -> str:
    """Build styles.xml with style 'ac' parameterized for 4-way differential.

    Defaults match 29dc6e's style "ac" exactly (full bug-repro).
    Each parameter can be overridden to "1" / "0" / different value to
    flip ONE candidate at a time.
    """
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="ac">
<w:name w:val="一太郎"/>
<w:pPr>
<w:widowControl w:val="0"/>
<w:wordWrap w:val="{word_wrap}"/>
<w:autoSpaceDE w:val="{auto_space_de}"/>
<w:autoSpaceDN w:val="{auto_space_dn}"/>
<w:adjustRightInd w:val="0"/>
<w:spacing w:line="210" w:lineRule="exact"/>
<w:jc w:val="both"/>
</w:pPr>
<w:rPr>
<w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="ＭＳ 明朝"/>
<w:spacing w:val="{char_spacing}"/>
<w:sz w:val="21"/><w:szCs w:val="21"/>
</w:rPr>
</w:style>
</w:styles>"""


SECT_PR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '</w:sectPr>'
)


def paragraph_xml(text: str, *, run_char_spacing: str = "0") -> str:
    """Paragraph using pStyle='ac', overriding line to 220 (11pt) and char
    spacing to 0 — matches 29dc6e i=265 exactly.
    """
    return (
        '<w:p>'
        '<w:pPr>'
        '<w:pStyle w:val="ac"/>'
        '<w:spacing w:before="120" w:line="220" w:lineRule="exact"/>'
        '<w:ind w:leftChars="150" w:left="543" '
        'w:rightChars="72" w:right="156" '
        'w:hangingChars="100" w:hanging="217"/>'
        f'<w:rPr><w:spacing w:val="{run_char_spacing}"/></w:rPr>'
        '</w:pPr>'
        '<w:r><w:rPr>'
        '<w:rFonts w:hint="eastAsia"/>'
        f'<w:spacing w:val="{run_char_spacing}"/>'
        '</w:rPr>'
        f'<w:t xml:space="preserve">{escape(text)}</w:t>'
        '</w:r>'
        '</w:p>'
    )


def cell_table_xml(text: str, *, run_char_spacing: str = "0") -> str:
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
        '</w:tblBorders>'
        '<w:tblLayout w:type="fixed"/>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="7541"/></w:tblGrid>'
        '<w:tr>'
        '<w:tc>'
        '<w:tcPr><w:tcW w:w="7541" w:type="dxa"/></w:tcPr>'
        + paragraph_xml(text, run_char_spacing=run_char_spacing) +
        '</w:tc>'
        '</w:tr>'
        '</w:tbl>'
    )


def write_docx(path: str, body_xml: str, styles: str) -> None:
    full = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '<w:body>\n'
        + body_xml + SECT_PR +
        '\n</w:body>\n</w:document>'
    )
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


# 29dc6e i=265 EXACT text
I265_TEXT = (
    "○　"
    "暴力団員等がその事業活動を支配する者又は暴力団員等をその"
    "業務に従事させ、若しくは当該業務の補助者として使用するおそれのある者"
)


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v4: FULL bug-repro baseline (style ac exactly as 29dc6e).
    # Expected: Word 3 lines, Oxi 2 lines (= bug repro'd).
    write_docx(
        os.path.join(OUT_DIR, "v4_style_ac_inherited.docx"),
        cell_table_xml(I265_TEXT, run_char_spacing="0"),
        styles_xml(),  # all defaults: wordWrap=0, autoSpaceDE=0, autoSpaceDN=0, char_spacing=-1
    )

    # v4a: Same as v4 but wordWrap=1 (Latin-no-mid-break).
    # If wordWrap is the differentiator, this should make Word wrap to 2 lines.
    write_docx(
        os.path.join(OUT_DIR, "v4a_wordwrap1.docx"),
        cell_table_xml(I265_TEXT, run_char_spacing="0"),
        styles_xml(word_wrap="1"),
    )

    # v4b: autoSpaceDE=1 + autoSpaceDN=1.
    # If autoSpaceDE/DN is the differentiator, this changes wrap.
    write_docx(
        os.path.join(OUT_DIR, "v4b_autospace1.docx"),
        cell_table_xml(I265_TEXT, run_char_spacing="0"),
        styles_xml(auto_space_de="1", auto_space_dn="1"),
    )

    # v4c: Strip 「、」punctuation (Kinsoku trigger candidates).
    text_no_comma = I265_TEXT.replace("、", "")
    write_docx(
        os.path.join(OUT_DIR, "v4c_no_kinsoku_comma.docx"),
        cell_table_xml(text_no_comma, run_char_spacing="0"),
        styles_xml(),
    )

    # v4d: char spacing at style level=0 (eliminates the -1 → 0 override question).
    write_docx(
        os.path.join(OUT_DIR, "v4d_charspacing0.docx"),
        cell_table_xml(I265_TEXT, run_char_spacing="0"),
        styles_xml(char_spacing="0"),
    )

    print("Done.")


if __name__ == "__main__":
    main()
