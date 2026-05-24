"""Author minimal repro fixtures isolating the `linesAndChars` docGrid
character-width behaviour observed in b837808d0555 P 11.

Each fixture has ONE paragraph of pure-CJK text plus margins/font matching
b837808d0555's section settings. Variables exercised across V0..V5:

  V0_baseline       — docGrid linesAndChars, linePitch=360, no charSpace,
                      MS Gothic, no explicit sz, no szCs override, no indent
  V1_szCs24         — V0 + <w:szCs w:val="24"/> (matches b837 P 11 run rPr)
  V2_indent         — V0 + <w:ind w:leftChars="100" w:left="240"/> (matches b837)
  V3_b837_exact     — V0 + szCs=24 + indent (matches b837 P 11 exactly)
  V4_linesGridOnly  — docGrid type="lines" linePitch=360 (no chars grid)
  V5_charSpaceZero  — V0 + explicit charSpace="0" (was: missing entirely)

All paragraphs use:
  - pgSz w=11906 h=16838 (A4 portrait)
  - pgMar top=1021 right=1418 bottom=1021 left=1418 (textArea=9070twips=453.5pt)
  - ＭＳ ゴシック (MS Gothic) east-Asian font
  - lang=ja-JP, no <w:sz> override (uses Word's default 10.5pt)

Body text in each: SAMPLE (96 CJK + few ASCII digit chars; long enough that
both Word and Oxi MUST wrap — exposes the wrap point as the diagnostic).

Output:
  tools/metrics/output/linesAndChars_repro/V0_baseline.docx
  ... V1, V2, V3, V4, V5

Run:
  python tools/metrics/build_linesAndChars_minrepro.py
"""
import os
import sys
import zipfile
from xml.sax.saxutils import escape

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(
    os.path.dirname(__file__), "output", "linesAndChars_repro"
)
os.makedirs(OUT_DIR, exist_ok=True)

# Reproduce P 11 of b837808d0555 sentence (the one with W=71ch, O=38ch in dml_diff).
# Long enough to wrap on either A4 layout regardless of width-per-char.
SAMPLE = (
    "我が国においては、平成23年３月11日の東日本大震災以降、"
    "政府、地方公共団体や事業者等が保有するデータの公開・"
    "活用に対する意識が高まった。"
)

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
  <w:rPrDefault>
    <w:rPr>
      <w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>
      <w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
    </w:rPr>
  </w:rPrDefault>
  <w:pPrDefault/>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a">
  <w:name w:val="Normal"/>
  <w:qFormat/>
  <w:pPr>
    <w:widowControl w:val="0"/>
    <w:jc w:val="both"/>
  </w:pPr>
  <w:rPr>
    <w:kern w:val="2"/>
    <w:sz w:val="24"/>
    <w:szCs w:val="22"/>
  </w:rPr>
</w:style>
</w:styles>"""

SECT_PR_GRID_LINESANDCHARS = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
    '<w:pgMar w:top="1021" w:right="1418" w:bottom="1021" w:left="1418" '
    'w:header="851" w:footer="397" w:gutter="0"/>'
    '<w:cols w:space="425"/>'
    '<w:docGrid w:type="linesAndChars" w:linePitch="360"/>'
    '</w:sectPr>'
)

SECT_PR_GRID_LINESONLY = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
    '<w:pgMar w:top="1021" w:right="1418" w:bottom="1021" w:left="1418" '
    'w:header="851" w:footer="397" w:gutter="0"/>'
    '<w:cols w:space="425"/>'
    '<w:docGrid w:type="lines" w:linePitch="360"/>'
    '</w:sectPr>'
)

SECT_PR_GRID_LINESANDCHARS_CHARSP0 = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
    '<w:pgMar w:top="1021" w:right="1418" w:bottom="1021" w:left="1418" '
    'w:header="851" w:footer="397" w:gutter="0"/>'
    '<w:cols w:space="425"/>'
    '<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="0"/>'
    '</w:sectPr>'
)


def build_document(sect_pr: str, pPr_extra: str, run_rpr_extra: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '<w:body>\n'
        '<w:p>'
        '<w:pPr>'
        f'{pPr_extra}'
        '<w:rPr>'
        '<w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:hint="eastAsia"/>'
        f'{run_rpr_extra}'
        '</w:rPr>'
        '</w:pPr>'
        '<w:r>'
        '<w:rPr>'
        '<w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:hint="eastAsia"/>'
        f'{run_rpr_extra}'
        '</w:rPr>'
        f'<w:t xml:space="preserve">{escape(SAMPLE)}</w:t>'
        '</w:r>'
        '</w:p>\n'
        f'{sect_pr}\n'
        '</w:body>\n'
        '</w:document>\n'
    )


def write_docx(out_path: str, doc_xml: str) -> None:
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
        zf.writestr("word/styles.xml", STYLES)
        zf.writestr("word/document.xml", doc_xml)


VARIANTS = [
    ("V0_baseline.docx",      SECT_PR_GRID_LINESANDCHARS,            "",                                            ""),
    ("V1_szCs24.docx",        SECT_PR_GRID_LINESANDCHARS,            "",                                            '<w:szCs w:val="24"/>'),
    ("V2_indent.docx",        SECT_PR_GRID_LINESANDCHARS,            '<w:ind w:leftChars="100" w:left="240"/>',     ""),
    ("V3_b837_exact.docx",    SECT_PR_GRID_LINESANDCHARS,            '<w:ind w:leftChars="100" w:left="240"/>',     '<w:szCs w:val="24"/>'),
    ("V4_linesGridOnly.docx", SECT_PR_GRID_LINESONLY,                "",                                            ""),
    ("V5_charSpaceZero.docx", SECT_PR_GRID_LINESANDCHARS_CHARSP0,    "",                                            ""),
]

for fname, sect, ppr, rpr in VARIANTS:
    xml = build_document(sect, ppr, rpr)
    out_path = os.path.join(OUT_DIR, fname)
    write_docx(out_path, xml)
    print(f"wrote {out_path}")

print(f"\nSample text ({len(SAMPLE)} chars): {SAMPLE!r}")
print(f"Total CJK + ASCII chars = {len(SAMPLE)}; expect wrap >=1 line at any plausible width.")
