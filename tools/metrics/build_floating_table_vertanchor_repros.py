"""Build minimal repro docx files isolating vertAnchor="text" floating
table footprint behavior.

Per [[session60-vertanchor-text-crossdoc-survey]] the 3a4f9f cascade
(score 0.1328) is caused by Oxi not reserving vertical footprint when a
`vertAnchor="text"` floating table renders. Before any layout fix we
need to confirm Word's actual placement rule via COM measurement on a
minimal isolated repro.

Variants (all on A4, top=99.25pt to match 3a4f9f's geometry):
  v1_small_y   : tblpY=22tw (≈1.1pt, mimics 3a4f9f doc1) + 3-row table
  v2_small_y_tall : tblpY=22tw + 10-row table (test footprint dependence on table height)
  v3_mid_y     : tblpY=300tw (15pt) + 3-row table (medium offset)
  v4_neg_y     : tblpY=-100tw (-5pt, mimics e201) + 3-row table (negative offset)
  v5_no_float  : same content but NO floating table (control for body cursor reference)

Each variant has:
  - 1 anchor paragraph "ANCHOR-PARAGRAPH-XYZ"
  - 1 vertAnchor="text" floating table (with rsidR for round-tripping)
  - 4 trailing body paragraphs "BODY-1", "BODY-2", "BODY-3", "BODY-4"
    so we can see where Word places each.

Output: c:/tmp/vfloat_v{1..5}.docx

Then: run `python tools/metrics/measure_floating_table_vertanchor_word.py`
to COM-measure Information(6) for ANCHOR / BODY-* in each.
"""
from __future__ import annotations

import os
import sys
import zipfile
from io import BytesIO

OUT_DIR = r"c:\tmp"

WORDPROCESSINGML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
</Types>
"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

WORD_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
</Relationships>
"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>
"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
"""

FONT_TABLE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
"""


def make_body_paragraph(text: str) -> str:
    return f"""<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"""


def make_inline_table(label: str, table_w: int = 4000) -> str:
    """Inline (non-floating) single-cell table containing one paragraph."""
    return f"""<w:tbl>
  <w:tblPr>
    <w:tblW w:w="{table_w}" w:type="dxa"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="{table_w}"/></w:tblGrid>
  <w:tr><w:tc>
    <w:tcPr><w:tcW w:w="{table_w}" w:type="dxa"/></w:tcPr>
    <w:p><w:r><w:t xml:space="preserve">{label}</w:t></w:r></w:p>
  </w:tc></w:tr>
</w:tbl>"""


def make_floating_table(tblpY: int, n_rows: int, table_w: int = 4000, horz_anchor: str | None = "margin", tblpX: int | None = None) -> str:
    """Build a single-column floating table with vertAnchor='text' and the given tblpY.

    `table_w` is the table width in twips. Default 4000tw ≈ 200pt (leaves wrap room
    on a 425pt content-width page). Pass ~8400tw to mimic 3a4f9f's full-width tables.
    `horz_anchor`: 'margin' (default), 'page', 'text', or None to OMIT the attribute
    (then Word defaults to 'column' — mimics ed025c's pattern).
    `tblpX`: horizontal offset in twips. None = omit (default 0 = flush at anchor).
    """
    rows_xml = []
    for i in range(n_rows):
        rows_xml.append(
            f"""<w:tr><w:tc>
  <w:tcPr><w:tcW w:w="{table_w}" w:type="dxa"/></w:tcPr>
  <w:p><w:r><w:t xml:space="preserve">CELL-{i + 1}</w:t></w:r></w:p>
</w:tc></w:tr>"""
        )
    rows = "\n".join(rows_xml)
    horz_attr = f' w:horzAnchor="{horz_anchor}"' if horz_anchor else ""
    tblpx_attr = f' w:tblpX="{tblpX}"' if tblpX is not None else ""
    return f"""<w:tbl>
  <w:tblPr>
    <w:tblpPr w:vertAnchor="text"{horz_attr}{tblpx_attr} w:tblpY="{tblpY}" w:leftFromText="142" w:rightFromText="142"/>
    <w:tblW w:w="{table_w}" w:type="dxa"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="{table_w}"/></w:tblGrid>
  {rows}
</w:tbl>"""


def build_document_xml(
    include_table: bool,
    tblpY: int,
    n_table_rows: int,
    n_body_after: int,
    table_w: int = 4000,
    horz_anchor: str | None = "margin",
    body_after_is_table: bool = False,
    tblpX: int | None = None,
) -> str:
    """Build the document.xml body.

    If body_after_is_table=True, the trailing content is N inline tables
    instead of body paragraphs (mimics ed025c's stacked-table pattern).
    """
    body = []
    body.append(make_body_paragraph("ANCHOR-PARAGRAPH-XYZ"))
    if include_table:
        body.append(make_floating_table(tblpY, n_table_rows, table_w, horz_anchor, tblpX))
    if body_after_is_table:
        # Stacked inline tables (NOT floating) — mimics ed025c
        for i in range(n_body_after):
            body.append(make_inline_table(f"AFTER-T{i + 1}", table_w))
    else:
        for i in range(n_body_after):
            body.append(make_body_paragraph(f"BODY-{i + 1}"))
    # 3a4f9f matching geometry: A4 (11906x16838), pgMar top=1985 bottom=1701
    sect_pr = """<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/>
  <w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>"""
    inner = "\n".join(body)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WORDPROCESSINGML_NS}">
<w:body>
{inner}
{sect_pr}
</w:body>
</w:document>"""


def write_docx(out_path: str, document_xml: str) -> None:
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        z.writestr("_rels/.rels", ROOT_RELS_XML)
        z.writestr("word/_rels/document.xml.rels", WORD_RELS_XML)
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
        z.writestr("word/fontTable.xml", FONT_TABLE_XML)


VARIANTS = [
    # Each tuple: (label, include_table, tblpY, n_table_rows, n_body_after, table_w, horz_anchor, body_after_is_table, tblpX)
    ("v1_small_y",       True,   22, 3,  4, 4000, "margin", False, None),
    ("v2_small_y_tall",  True,   22, 10, 4, 4000, "margin", False, None),
    ("v3_mid_y",         True,  300, 3,  4, 4000, "margin", False, None),
    ("v4_neg_y",         True, -100, 3,  4, 4000, "margin", False, None),
    ("v5_no_float",      False,   0, 0,  4, 0,    "margin", False, None),
    # Full-width on 425pt content_w (ratio 0.99)
    ("v6_fullw_small_y", True,   22, 3,  4, 8400, "margin", False, None),
    ("v7_fullw_tall",    True,   22, 10, 4, 8400, "margin", False, None),
    ("v8_fullw_mid_y",   True,  300, 3,  4, 8400, "margin", False, None),
    # Differentiator tests for ed025c pattern
    ("v9_fullw_no_horz",          True,   22, 3, 4, 8400, None,     False, None),
    ("v10_fullw_horz_column",     True,   22, 3, 4, 8400, "column", False, None),
    ("v11_fullw_body_is_table",   True,   22, 3, 4, 8400, "margin", True,  None),
    ("v12_fullw_no_horz_body_tbl",True,   22, 3, 4, 8400, None,     True,  None),
    # NEW (session 60 part 3): tblpX != 0 — mimic ed025c's positioned tables.
    # ed025c has tblpX=641 (32pt), 817 (40pt), 2008 (100pt) with full-width tables.
    ("v13_fullw_horz_missing_tblpX641", True, 22, 3, 4, 8400, None,     False, 641),  # most common ed025c pattern
    ("v14_fullw_horz_margin_tblpX641",  True, 22, 3, 4, 8400, "margin", False, 641),  # control: same tblpX but margin anchor
    ("v15_fullw_horz_page_tblpX2008",   True, 22, 3, 4, 8400, "page",   False, 2008), # ed025c table 7 pattern
    ("v16_narrow_horz_missing_tblpX641",True, 22, 3, 4, 4000, None,     False, 641),  # narrow + ed025c-like position
]


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)
    print(f"=== Building minimal repros for vertAnchor=\"text\" footprint study ===")
    for label, include_table, tblpY, n_rows, n_body, table_w, horz_anchor, body_is_tbl, tblpX in VARIANTS:
        xml = build_document_xml(include_table, tblpY, n_rows, n_body, table_w, horz_anchor, body_is_tbl, tblpX)
        out = os.path.join(OUT_DIR, f"vfloat_{label}.docx")
        write_docx(out, xml)
        print(f"  {label:<38} tblpY={tblpY:>5}  n_rows={n_rows:>2}  table_w={table_w:>5}  horz={str(horz_anchor):<8}  tblpX={str(tblpX):<5}  body_tbl={int(body_is_tbl)}")
    print(f"\nTotal: {len(VARIANTS)} variants written to {OUT_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
