"""Author bug-REPRODUCING minimal repro for the S299 Stage 3 hypothesis.

S299 Stage 3 hypothesis: the unexplained +11pt gap between 29dc6e i=265
and i=266 in Word (vs Oxi which has no excess) is driven by i=266's
unique `w:afterLines="50" w:after="146"` properties — possibly Word
applies both cumulatively, or attributes one of them as BEFORE-NEXT
rather than AFTER-CURRENT.

Where the prior repro (`build_phase2_wrap_repro.py`) showed Oxi handles
the wrap CORRECTLY for the i=265 text in isolation (3 lines at sz=22,
or 2 lines at sz=21 matching Word), this repro isolates the SPACING
interaction:

  Cell with TWO paragraphs:
    Paragraph 1: exact 11pt line, ~64 CJK chars wrapping to 2 lines
    Paragraph 2: same line config + UNIQUELY adds w:afterLines/w:after

Measurement: gap between paragraph 1 START y and paragraph 2 START y.
  Expected if hypothesis is FALSIFIED (Oxi correct):  2*11 + 6 = 28pt
  Expected if hypothesis is CONFIRMED (Word excess):  ~39pt (+11 unaccounted)

If Oxi gives 28pt AND Word gives 39pt on this minimal repro → root
cause confirmed in isolation. Next: implement fix.
If both give same value → Stage 3 also falsified, re-derive.

Uses sz=21 (matches 29dc6e style "ac" font size) so wrap counts
match between Word and Oxi as established in S299.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(
    os.path.dirname(__file__), "..", "fixtures", "phase2_afterlines_samples"
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

# Use sz=21 (10.5pt) to match 29dc6e style "ac" — at this size, V1's 64-char
# text wraps to 2 lines in BOTH Word and Oxi, eliminating the wrap-count confound.
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝"/>
<w:sz w:val="21"/><w:szCs w:val="21"/>
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


def _paragraph_xml(text: str, with_afterlines: bool) -> str:
    """One paragraph with exact 11pt line, hanging indent, sz=21.

    If with_afterlines, additionally include w:afterLines and w:after
    (matching 29dc6e i=266's UNIQUE properties vs i=265).
    """
    afterlines_attrs = ' w:afterLines="50" w:after="146"' if with_afterlines else ''
    return (
        '<w:p>'
        '<w:pPr>'
        f'<w:spacing w:before="120"{afterlines_attrs} w:line="220" w:lineRule="exact"/>'
        '<w:ind w:left="543" w:right="156" w:hanging="217"/>'
        '<w:rPr><w:spacing w:val="0"/></w:rPr>'
        '</w:pPr>'
        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:spacing w:val="0"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(text)}</w:t>'
        '</w:r>'
        '</w:p>'
    )


def cell_table_xml(text1: str, text2: str, p2_has_afterlines: bool) -> str:
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
        + _paragraph_xml(text1, with_afterlines=False) +
        _paragraph_xml(text2, with_afterlines=p2_has_afterlines) +
        '</w:tc>'
        '</w:tr>'
        '</w:tbl>'
    )


def write_docx(path: str, body_xml: str) -> None:
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


# Paragraph 1 text: 29dc6e i=265 (wraps to 2 lines at sz=21)
P1_TEXT = (
    "○　"
    "暴力団員等がその事業活動を支配する者又は暴力団員等をその"
    "業務に従事させ若しくは当該業務の補助者として使用するおそれのある者"
)
# Paragraph 2 text: 29dc6e i=266
P2_TEXT = (
    "○　"
    "統計法令に基づく罰則の適用を受けている者、調査票情報又は匿名データ"
    "を利用して不適切な行為を行った者"
)


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")
    # Control: both paragraphs WITHOUT afterLines+after.
    # Expected gap: 2*11 + 6 = 28pt.
    write_docx(
        os.path.join(OUT_DIR, "control_no_afterlines.docx"),
        cell_table_xml(P1_TEXT, P2_TEXT, p2_has_afterlines=False),
    )
    # Test: P2 has afterLines=50 + after=146 (matching 29dc6e i=266).
    # If Word's gap > 28pt here while Oxi's is still 28pt → bug repro'd.
    write_docx(
        os.path.join(OUT_DIR, "p2_afterlines_after.docx"),
        cell_table_xml(P1_TEXT, P2_TEXT, p2_has_afterlines=True),
    )
    print("Done.")
    print()
    print(f"P1 text len: {len(P1_TEXT)}")
    print(f"P2 text len: {len(P2_TEXT)}")


if __name__ == "__main__":
    main()
