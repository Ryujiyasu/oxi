"""Author minimal w:ruby repro fixtures for COM measurement.

Generates 4 fixtures with progressively more ruby surface area:

  RUBY_V1_basic.docx        — single short ruby (1 base char + 2 ruby chars), no rubyPr
  RUBY_V2_align_variants.docx — 5 paragraphs, one per w:rubyAlign value
  RUBY_V3_hps_variants.docx  — 3 paragraphs varying hps + hpsRaise
  RUBY_V4_lineheight.docx    — paragraph w/ ruby vs paragraph w/o ruby, same font/size

All paragraphs use:
  - Page: A4 portrait, 25.4mm margins (Word default)
  - Font: ＭＳ 明朝 (MS Mincho) 10.5pt (== sz val=21)
  - lang=ja-JP, no explicit lineSpacing override
  - sectPr: docGrid type=default linePitch=312

This produces real <w:ruby><w:rubyPr/><w:rt><w:r/></w:rt><w:rubyBase><w:r/></w:rubyBase></w:ruby>
markup that Word recognizes and renders with full furigana semantics.

Usage:
  python tools/metrics/build_ruby_repro_fixtures.py
Output:
  pipeline_data/docx/RUBY_V1_basic.docx
  pipeline_data/docx/RUBY_V2_align_variants.docx
  pipeline_data/docx/RUBY_V3_hps_variants.docx
  pipeline_data/docx/RUBY_V4_lineheight.docx
"""
import os
import sys
import zipfile
from xml.sax.saxutils import escape

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = "pipeline_data/docx"

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

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault>
<w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>
<w:sz w:val="21"/>
<w:szCs w:val="21"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr>
</w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
</w:styles>"""

def _sect_pr(line_pitch: int = 312, grid_type: str = "default") -> str:
    return (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:cols w:space="425"/>'
        f'<w:docGrid w:type="{grid_type}" w:linePitch="{line_pitch}"/>'
        '</w:sectPr>'
    )

SECT_PR = _sect_pr(312)

DOC_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
"""
DOC_TAIL_DEFAULT = SECT_PR + """
</w:body>
</w:document>"""


def _r(text: str, sz: str = "21") -> str:
    """Plain run with given size (half-points)."""
    return (
        f'<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" '
        f'w:hAnsi="ＭＳ 明朝"/><w:sz w:val="{sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
    )


def _ruby(base: str, ruby_text: str, *, ruby_align: str = None,
          hps: str = None, hps_raise: str = None, hps_base_text: str = None,
          base_sz: str = "21", ruby_sz: str = "11") -> str:
    """Build a w:ruby element. base_sz/ruby_sz are half-points (21 = 10.5pt)."""
    parts = ["<w:rubyPr>"]
    if ruby_align:
        parts.append(f'<w:rubyAlign w:val="{ruby_align}"/>')
    if hps:
        parts.append(f'<w:hps w:val="{hps}"/>')
    if hps_raise:
        parts.append(f'<w:hpsRaise w:val="{hps_raise}"/>')
    if hps_base_text:
        parts.append(f'<w:hpsBaseText w:val="{hps_base_text}"/>')
    parts.append('<w:lid w:val="ja-JP"/>')
    parts.append("</w:rubyPr>")
    rubypr_xml = "".join(parts)
    rt_run = (
        f'<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" '
        f'w:hAnsi="ＭＳ 明朝"/><w:sz w:val="{ruby_sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(ruby_text)}</w:t></w:r>'
    )
    base_run = (
        f'<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" '
        f'w:hAnsi="ＭＳ 明朝"/><w:sz w:val="{base_sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(base)}</w:t></w:r>'
    )
    return (
        f"<w:r><w:ruby>{rubypr_xml}"
        f"<w:rt>{rt_run}</w:rt>"
        f"<w:rubyBase>{base_run}</w:rubyBase>"
        f"</w:ruby></w:r>"
    )


def _para(*runs: str) -> str:
    inner = "".join(runs)
    return f"<w:p>{inner}</w:p>"


def write_docx(path: str, paragraphs_xml: str, *, line_pitch: int = 312, grid_type: str = "default") -> None:
    sect_pr = _sect_pr(line_pitch, grid_type)
    body_xml = DOC_HEAD + paragraphs_xml + sect_pr + "\n</w:body>\n</w:document>"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", body_xml)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


def build_v1_basic() -> None:
    """Single ruby: 漢[かん] in a one-line paragraph."""
    paragraphs = "\n".join([
        _para(
            _r("これは"),
            _ruby("漢字", "かんじ"),
            _r("のテストです。"),
        ),
    ])
    write_docx(os.path.join(OUT_DIR, "RUBY_V1_basic.docx"), paragraphs)


def build_v2_align_variants() -> None:
    """One paragraph per w:rubyAlign value. ECMA-376 §17.3.3.25 lists:
    center, distributeLetter, distributeSpace, left, right, rightVertical.
    """
    aligns = ["center", "distributeLetter", "distributeSpace", "left", "right"]
    paragraphs = []
    for a in aligns:
        paragraphs.append(_para(
            _r(f"[{a}]: "),
            _ruby("特定", "とくてい", ruby_align=a),
            _r("の単語にルビ"),
        ))
    write_docx(os.path.join(OUT_DIR, "RUBY_V2_align_variants.docx"), "\n".join(paragraphs))


def build_v3_hps_variants() -> None:
    """Vary hps (ruby font size in half-pt) and hpsRaise (raise above baseline in half-pt).
    Default ruby_sz=11 = 5.5pt (half of 10.5pt base). Try 9 (4.5pt), 13 (6.5pt) too."""
    paragraphs = [
        _para(
            _r("[hps=11 (default 5.5pt)] "),
            _ruby("漢字", "かんじ", hps="11", hps_base_text="21"),
        ),
        _para(
            _r("[hps=9 (4.5pt smaller)] "),
            _ruby("漢字", "かんじ", hps="9", hps_base_text="21", ruby_sz="9"),
        ),
        _para(
            _r("[hps=13 (6.5pt larger)] "),
            _ruby("漢字", "かんじ", hps="13", hps_base_text="21", ruby_sz="13"),
        ),
        _para(
            _r("[hpsRaise=18, hps=11] "),
            _ruby("漢字", "かんじ", hps="11", hps_raise="18", hps_base_text="21"),
        ),
        _para(
            _r("[hpsRaise=24, hps=11] "),
            _ruby("漢字", "かんじ", hps="11", hps_raise="24", hps_base_text="21"),
        ),
    ]
    write_docx(os.path.join(OUT_DIR, "RUBY_V3_hps_variants.docx"), "\n".join(paragraphs))


def build_v4_lineheight() -> None:
    """3 paragraphs to isolate line-height effect:
    P1 (no ruby) — baseline reference for line-height
    P2 (with ruby) — to compare paragraph-Y offset to P3
    P3 (no ruby) — baseline reference (after ruby paragraph)
    All paragraphs identical font/size; the ONLY variable is presence of <w:ruby>.
    """
    paragraphs = [
        _para(_r("ルビなし段落１: 通常の本文行のベースライン Y を確認します。")),
        _para(
            _r("ルビあり段落: ふりがなを"),
            _ruby("含", "ふく"),
            _r("む段落の Y 位置と次行への影響を測定。"),
        ),
        _para(_r("ルビなし段落２: ルビ段落の直後に置いて、累積効果を測定します。")),
    ]
    write_docx(os.path.join(OUT_DIR, "RUBY_V4_lineheight.docx"), "\n".join(paragraphs))


def build_v5_linepitch_variants() -> None:
    """Test whether ruby expansion is absorbed at larger linePitch (inter-line gap).

    Per ra_manual_measurements 'ruby_no_line_height_impact: true (in most cases)'
    — what are 'most cases'? Hypothesis: when linePitch_pt - no_ruby_line_height
    >= ruby_expansion, the gap absorbs the ruby and dy stays at no-ruby value.

    Builds 5 separate docs (one per linePitch). Each contains the V4 triple
    structure (no-ruby / with-ruby / no-ruby) so we can compare ruby vs
    baseline dy at each linePitch.
    """
    pitches = [240, 280, 312, 360, 400, 480]
    for lp in pitches:
        paragraphs = [
            _para(_r(f"linePitch={lp} no-ruby P1.")),
            _para(
                _r("with-ruby P2: "),
                _ruby("含", "ふく"),
                _r("む段落です。"),
            ),
            _para(_r(f"linePitch={lp} no-ruby P3.")),
        ]
        write_docx(
            os.path.join(OUT_DIR, f"RUBY_V5_linepitch_{lp}.docx"),
            "\n".join(paragraphs),
            line_pitch=lp,
        )


def build_v5b_linepitch_lines_grid() -> None:
    """Re-test linePitch absorption with docGrid type='lines' (V5 with type='default' was invalid).
    Each doc: V4 structure (no-ruby/ruby/no-ruby). Compare dy to detect whether large
    linePitch absorbs the ruby expansion.
    """
    pitches = [240, 280, 312, 360, 400, 480]
    for lp in pitches:
        paragraphs = [
            _para(_r(f"linePitch={lp} type=lines no-ruby P1.")),
            _para(
                _r("with-ruby P2: "),
                _ruby("含", "ふく"),
                _r("む段落です。"),
            ),
            _para(_r(f"linePitch={lp} type=lines no-ruby P3.")),
        ]
        write_docx(
            os.path.join(OUT_DIR, f"RUBY_V5b_lines_{lp}.docx"),
            "\n".join(paragraphs),
            line_pitch=lp,
            grid_type="lines",
        )


def build_v6_hpsRaise_variants() -> None:
    """Derive the hpsRaise → expansion rule. V3 gave only 1 raise observation.
    Tests raise ∈ {0, 6, 12, 18, 24, 36, 48} (half-pt) all with hps=11.

    Each paragraph has ruby with same hps=11 but differing hpsRaise. dy from
    paragraph N to N+1 reveals paragraph N's line height (per V3 method).
    """
    raises = [None, 6, 12, 18, 24, 36, 48]
    paragraphs = []
    for r in raises:
        label = "hpsRaise=None" if r is None else f"hpsRaise={r}"
        paragraphs.append(_para(
            _r(f"[{label}]: "),
            _ruby(
                "漢字", "かんじ",
                hps="11",
                hps_raise=str(r) if r is not None else None,
                hps_base_text="21",
            ),
            _r("テスト"),
        ))
    write_docx(os.path.join(OUT_DIR, "RUBY_V6_hpsRaise.docx"), "\n".join(paragraphs))


def build_v7_wrap_interaction() -> None:
    """Long paragraph with ruby placed near the line-end position to measure
    whether base+ruby acts as a single wrap unit, or whether the ruby base
    can wrap mid-element.

    Strategy: pad paragraph so that the ruby's base text would naturally
    cross the right margin if treated as ordinary text. If Word treats
    the ruby as atomic, the entire ruby element wraps to next line.
    """
    paragraphs = [
        _para(
            _r("これは折り返し試験段落です。" * 3),  # 30+ chars
            _ruby("漢字練習", "かんじれんしゅう"),
            _r("で行末ルビ折返動作を確認"),
        ),
        _para(
            _r("ルビなし長段落の比較対照: " * 4),
        ),
    ]
    write_docx(os.path.join(OUT_DIR, "RUBY_V7_wrap.docx"), "\n".join(paragraphs))


def build_v12_atomic_wrap_overhang() -> None:
    """Test that ruby_w > base_w forces wrap budget reduction.

    Constructs a paragraph that, with ruby NOT considered, fits exactly N
    chars per line. With the ruby's 1-2pt overhang reserved, the line
    should wrap one char earlier.

    Pattern: 40+ fullwidth chars at 10.5pt = 420pt+, near 432pt body width.
    A ruby-bearing run "特定" + "とくてい" (1pt overhang) is placed near the
    line end. If atomic wrap reservation works, the line breaks 1 char
    before where it would without ruby consideration.
    """
    # Header that pads close to wrap boundary
    pad_text = "あ" * 38  # 38 × 10.5 = 399pt
    paragraphs = [
        # P1: 38 chars + ruby base (2 chars=21pt) + ruby (4chars=22pt) + 2 chars = total
        # base reading: 38+2 = 40 fullwidth + 2 more = 42*10.5 = 441 > 432 wrap
        # but with overhang reserved: should break around char 39-40
        _para(
            _r(pad_text),
            _ruby("特定", "とくてい", hps="11"),
            _r("はパッド"),
        ),
        _para(_r("ルビなし対比: " + "あ" * 40)),
    ]
    write_docx(os.path.join(OUT_DIR, "RUBY_V12_atomic_wrap.docx"), "\n".join(paragraphs))


def build_v11_align_asymmetric() -> None:
    """Test rubyAlign with ASYMMETRIC base/ruby widths (base_w != ruby_w).

    V2 used base="特定" + ruby="とくてい" with hps causing equal widths
    (21pt each), so all align modes produced identical positioning.

    V11 uses "漢字" (2 chars × 10.5 = 21pt) with "かん" (2 chars × 5.5 =
    11pt) — base wider than ruby — so center/left/right/distributeLetter/
    distributeSpace produce visibly distinct X positions.
    """
    aligns = ["center", "distributeLetter", "distributeSpace", "left", "right"]
    paragraphs = []
    for a in aligns:
        paragraphs.append(_para(
            _r(f"[{a}]: "),
            _ruby("漢字", "かん", ruby_align=a, hps="11"),
            _r("でテスト"),
        ))
    write_docx(os.path.join(OUT_DIR, "RUBY_V11_align_asymmetric.docx"), "\n".join(paragraphs))


def build_v10_base_size_variants() -> None:
    """Test whether the empirical formula
       expansion = max(0, hpsRaise_pt + 0.75 * hps_pt - line_box_ascent)
    holds at non-10.5pt base sizes.

    For 10.5pt base: line_box_ascent = ~9pt observed.
    Hypothesis: line_box_ascent = base_pt * 0.857 (font-typical ascent ratio).

    Predictions for default-hps default-raise at each base size:
        base=9pt:  default_hps=4.5pt, default_raise=base*0.857=7.71pt
                   expansion = max(0, 7.71 + 0.75*4.5 - 7.71) = 3.375pt
        base=12pt: default_hps=6pt,   default_raise=10.29pt
                   expansion = max(0, 10.29 + 0.75*6 - 10.29) = 4.5pt
        base=14pt: default_hps=7pt,   default_raise=12pt
                   expansion = max(0, 12 + 0.75*7 - 12) = 5.25pt

    Generates 3 separate docs with V4 structure (no-ruby/ruby/no-ruby) at
    each base size. dy measurement reveals base-specific no-ruby LH and
    ruby LH.
    """
    base_sizes_halfpt = [18, 24, 28]  # 9pt, 12pt, 14pt
    for base_sz in base_sizes_halfpt:
        base_pt = base_sz / 2
        # body run with this base size
        body_paras = [
            _para(_r(f"base={base_pt}pt no-ruby P1.", sz=str(base_sz))),
            _para(
                _r("with-ruby P2: ", sz=str(base_sz)),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(base_sz // 2),  # default ratio
                ),
                _r("む段落です。", sz=str(base_sz)),
            ),
            _para(_r(f"base={base_pt}pt no-ruby P3.", sz=str(base_sz))),
        ]
        write_docx(
            os.path.join(OUT_DIR, f"RUBY_V10_base_{int(base_pt*10):03d}dpt.docx"),
            "\n".join(body_paras),
        )


def build_v9_combined_formula_grid() -> None:
    """Cross-validate the combined formula
        ruby_expansion_pt = max(hps_pt - 1.5, hpsRaise_pt - 5)
    by testing hps in {9, 13} (=4.5pt, 6.5pt) with explicit hpsRaise variants.

    For hps=9 (4.5pt), expansion from hps alone = 4.5 - 1.5 = 3.0pt.
        - raise=6  (3pt):  max(3.0, max(0, 3-5))  = max(3.0, 0)  = 3.0
        - raise=12 (6pt):  max(3.0, max(0, 6-5))  = max(3.0, 1)  = 3.0
        - raise=18 (9pt):  max(3.0, max(0, 9-5))  = max(3.0, 4)  = 4.0
        - raise=24 (12pt): max(3.0, max(0, 12-5)) = max(3.0, 7)  = 7.0
        - raise=36 (18pt): max(3.0, max(0, 18-5)) = max(3.0, 13) = 13.0

    For hps=13 (6.5pt), expansion from hps alone = 6.5 - 1.5 = 5.0pt.
        - raise=6  (3pt):  max(5.0, 0)  = 5.0
        - raise=12 (6pt):  max(5.0, 1)  = 5.0
        - raise=18 (9pt):  max(5.0, 4)  = 5.0
        - raise=24 (12pt): max(5.0, 7)  = 7.0
        - raise=36 (18pt): max(5.0, 13) = 13.0
    """
    paragraphs = []
    grid = [
        # (hps_halfpt, raise_halfpt, predicted_pt, label_tag)
        (9,  6,  3.0,  "hps=9 raise=6"),
        (9,  12, 3.0,  "hps=9 raise=12"),
        (9,  18, 4.0,  "hps=9 raise=18"),
        (9,  24, 7.0,  "hps=9 raise=24"),
        (9,  36, 13.0, "hps=9 raise=36"),
        (13, 6,  5.0,  "hps=13 raise=6"),
        (13, 12, 5.0,  "hps=13 raise=12"),
        (13, 18, 5.0,  "hps=13 raise=18"),
        (13, 24, 7.0,  "hps=13 raise=24"),
        (13, 36, 13.0, "hps=13 raise=36"),
    ]
    for hps, raise_, _pred, label in grid:
        paragraphs.append(_para(
            _r(f"[{label}]: "),
            _ruby(
                "漢字", "かんじ",
                hps=str(hps),
                hps_raise=str(raise_),
                hps_base_text="21",
                ruby_sz=str(hps),
            ),
            _r("確認"),
        ))
    write_docx(os.path.join(OUT_DIR, "RUBY_V9_combined_grid.docx"), "\n".join(paragraphs))


def build_v8_extreme_hps() -> None:
    """Test whether ruby_expansion = hps_pt - 1.5 holds OUTSIDE the V3 range
    (9-13 half-pt). Tests hps ∈ {5, 7, 15, 17, 21} half-pt with base=21.

    Predictions per linear fit:
      hps=5  (2.5pt) → 1.0pt expansion
      hps=7  (3.5pt) → 2.0pt expansion
      hps=15 (7.5pt) → 6.0pt expansion
      hps=17 (8.5pt) → 7.0pt expansion
      hps=21 (10.5pt= base) → 9.0pt expansion
    """
    hps_values = [5, 7, 15, 17, 21]
    paragraphs = []
    for h in hps_values:
        paragraphs.append(_para(
            _r(f"[hps={h}]: "),
            _ruby(
                "漢字", "かんじ",
                hps=str(h),
                hps_base_text="21",
                ruby_sz=str(h),
            ),
            _r("確認"),
        ))
    write_docx(os.path.join(OUT_DIR, "RUBY_V8_extreme_hps.docx"), "\n".join(paragraphs))


def main() -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    print(f"Writing fixtures to {OUT_DIR}/")
    build_v1_basic()
    build_v2_align_variants()
    build_v3_hps_variants()
    build_v4_lineheight()
    build_v5_linepitch_variants()
    build_v5b_linepitch_lines_grid()
    build_v6_hpsRaise_variants()
    build_v7_wrap_interaction()
    build_v8_extreme_hps()
    build_v9_combined_formula_grid()
    build_v10_base_size_variants()
    build_v11_align_asymmetric()
    build_v12_atomic_wrap_overhang()
    print("Done.")


if __name__ == "__main__":
    main()
