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


def _r(text: str, sz: str = "21", font_name: str = "ＭＳ 明朝") -> str:
    """Plain run with given size (half-points). font_name controls all 3 rFonts attrs."""
    return (
        f'<w:r><w:rPr><w:rFonts w:ascii="{font_name}" w:eastAsia="{font_name}" '
        f'w:hAnsi="{font_name}"/><w:sz w:val="{sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
    )


def _ruby(base: str, ruby_text: str, *, ruby_align: str = None,
          hps: str = None, hps_raise: str = None, hps_base_text: str = None,
          base_sz: str = "21", ruby_sz: str = "11",
          font_name: str = "ＭＳ 明朝") -> str:
    """Build a w:ruby element. base_sz/ruby_sz are half-points (21 = 10.5pt).
    font_name controls all 3 rFonts attrs on both rt and rubyBase."""
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
        f'<w:r><w:rPr><w:rFonts w:ascii="{font_name}" w:eastAsia="{font_name}" '
        f'w:hAnsi="{font_name}"/><w:sz w:val="{ruby_sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(ruby_text)}</w:t></w:r>'
    )
    base_run = (
        f'<w:r><w:rPr><w:rFonts w:ascii="{font_name}" w:eastAsia="{font_name}" '
        f'w:hAnsi="{font_name}"/><w:sz w:val="{base_sz}"/></w:rPr>'
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


def build_v13_base_raise_grid() -> None:
    """Disambiguate base × hpsRaise × hps scaling for non-10.5pt bases.

    V10 already showed `expansion = max(0, raise + 0.75*hps - 9)` does NOT
    generalize beyond 10.5pt — at base=14pt with default ruby it was off
    by +1.25pt. The constant `9` was the MS Mincho line-box ascent at 10.5pt.
    Hypotheses to discriminate at base ∈ {9, 11, 12, 14}pt:
      H1 (font-ratio):   ascent = base_pt × (9/10.5) = base × 0.857
      H2 (constant):     ascent = 9pt (independent of base)
      H3 (CJK 9/14):     ascent = base_pt × 9/14 ≈ 0.643 × base
      H4 (CJK 9/7):      ascent = base_pt × 9/7 = no_ruby_LH itself

    For each base size, emit 9 paragraphs in V4-triple chain:
      P0 no-ruby ref
      P1 default ruby            (raise=default, hps=base/2)
      P2 no-ruby ref             ← gives P1 height
      P3 explicit raise=12halfpt (=6pt), hps=default
      P4 no-ruby ref
      P5 explicit raise=24halfpt (=12pt), hps=default
      P6 no-ruby ref
      P7 explicit hps=base_halfpt (=base_pt), raise=default
      P8 no-ruby closure ref

    dy from no-ruby anchor → ruby para → no-ruby reveals each cell's LH.
    Solving across cells:
      P1 cell    → default expansion = default_raise + 0.75 × (base/2) − ascent
      P3,P5 cell → ascent = (raise + 0.75 × default_hps) − measured_expansion
      P7 cell    → big_hps expansion = default_raise + 0.75 × base − ascent
    Two unknowns (default_raise, ascent) from 4 equations → over-determined.
    """
    base_sizes_halfpt = [18, 22, 24, 28]  # 9pt, 11pt, 12pt, 14pt
    for base_sz in base_sizes_halfpt:
        base_pt = base_sz / 2
        default_hps_halfpt = base_sz // 2  # base/2 in halfpt
        big_hps_halfpt = base_sz            # = base_pt in halfpt
        body_paras = [
            _para(_r(f"V13 base={base_pt}pt P0 no-ruby ref.", sz=str(base_sz))),
            _para(
                _r("P1 default-ruby: ", sz=str(base_sz)),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(default_hps_halfpt),
                    hps=str(default_hps_halfpt),
                ),
                _r("です。", sz=str(base_sz)),
            ),
            _para(_r("P2 no-ruby (measures P1 height).", sz=str(base_sz))),
            _para(
                _r("P3 raise=12halfpt: ", sz=str(base_sz)),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(default_hps_halfpt),
                    hps=str(default_hps_halfpt),
                    hps_raise="12",
                ),
                _r("です。", sz=str(base_sz)),
            ),
            _para(_r("P4 no-ruby (measures P3 height).", sz=str(base_sz))),
            _para(
                _r("P5 raise=24halfpt: ", sz=str(base_sz)),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(default_hps_halfpt),
                    hps=str(default_hps_halfpt),
                    hps_raise="24",
                ),
                _r("です。", sz=str(base_sz)),
            ),
            _para(_r("P6 no-ruby (measures P5 height).", sz=str(base_sz))),
            _para(
                _r(f"P7 hps={big_hps_halfpt}halfpt: ", sz=str(base_sz)),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(big_hps_halfpt),
                    hps=str(big_hps_halfpt),
                ),
                _r("です。", sz=str(base_sz)),
            ),
            _para(_r("P8 no-ruby closure (measures P7 height).", sz=str(base_sz))),
        ]
        write_docx(
            os.path.join(OUT_DIR, f"RUBY_V13_base_{int(base_pt*10):03d}dpt.docx"),
            "\n".join(body_paras),
        )


def build_v14_font_family_variants() -> None:
    """Test ROUND 10 PREDICTION: ascent constant = base × OS/2.sTypoAscender / unitsPerEm.

    TTF probe (round 10) found two main JP font families:
      MS legacy   (upem=256, ratio=0.8594) → 14pt ascent = 12.031pt
      Yu/BIZ std  (upem=2048, ratio=0.8799) → 14pt ascent = 12.319pt (Δ +0.288pt)

    With the V13-confirmed formula
      expansion = max(0, raise + 0.75×hps - base × ratio)
    at base=14pt, raise=12pt (hps=24halfpt), hps=7pt (default base/2):
      MS Mincho:  max(0, 12 + 5.25 - 12.031) = 5.219pt
      Yu Mincho:  max(0, 12 + 5.25 - 12.319) = 4.931pt
      Δ = 0.288pt → with Word's 0.5pt rounding, MS likely rounds to 5.5,
      Yu likely rounds to 5.0 → DETECTABLE.

    For each test font, emit 9 paragraphs in V13's V4-triple chain pattern,
    all base=14pt:
      P1 no-ruby ref
      P2 default ruby
      P3 no-ruby
      P4 explicit raise=12halfpt (=6pt), hps=default
      P5 no-ruby
      P6 explicit raise=24halfpt (=12pt), hps=default  ← signal cell
      P7 no-ruby
      P8 explicit hps=base_halfpt (=14pt), raise=default
      P9 no-ruby closure

    If V14 measured cells match TTF prediction within ±0.3pt rounding, the
    font-intrinsic generalization (round 10) is confirmed.
    """
    # (font_name_in_xml, file_suffix, predicted_family)
    fonts = [
        ("ＭＳ 明朝", "MSMincho_control", "MS_legacy_0.8594"),
        ("Yu Mincho", "YuMincho", "Yu_BIZ_std_0.8799"),
        ("游ゴシック", "YuGothic_jp", "Yu_BIZ_std_0.8799"),
        ("Yu Gothic", "YuGothic_en", "Yu_BIZ_std_0.8799"),
        ("BIZ UDMincho Medium", "BIZUDMincho", "Yu_BIZ_std_0.8799"),
    ]
    base_sz = 28  # halfpt = 14pt
    base_pt = 14.0
    default_hps = base_sz // 2  # halfpt
    big_hps = base_sz             # halfpt = base_pt
    for font_name, suffix, family in fonts:
        body_paras = [
            _para(_r(f"V14 P0 ref.", sz=str(base_sz), font_name=font_name)),
            _para(
                _r("P1 def: ", sz=str(base_sz), font_name=font_name),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(default_hps),
                    hps=str(default_hps),
                    font_name=font_name,
                ),
                _r("です。", sz=str(base_sz), font_name=font_name),
            ),
            _para(_r("V14 P2 ref.", sz=str(base_sz), font_name=font_name)),
            _para(
                _r("P3 r6: ", sz=str(base_sz), font_name=font_name),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(default_hps),
                    hps=str(default_hps),
                    hps_raise="12",
                    font_name=font_name,
                ),
                _r("です。", sz=str(base_sz), font_name=font_name),
            ),
            _para(_r("V14 P4 ref.", sz=str(base_sz), font_name=font_name)),
            _para(
                _r("P5 r12: ", sz=str(base_sz), font_name=font_name),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(default_hps),
                    hps=str(default_hps),
                    hps_raise="24",
                    font_name=font_name,
                ),
                _r("です。", sz=str(base_sz), font_name=font_name),
            ),
            _para(_r("V14 P6 ref.", sz=str(base_sz), font_name=font_name)),
            _para(
                _r(f"P7 hpsB: ", sz=str(base_sz), font_name=font_name),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(big_hps),
                    hps=str(big_hps),
                    font_name=font_name,
                ),
                _r("です。", sz=str(base_sz), font_name=font_name),
            ),
            _para(_r("V14 P8 close.", sz=str(base_sz), font_name=font_name)),
        ]
        write_docx(
            os.path.join(OUT_DIR, f"RUBY_V14_{suffix}_140dpt.docx"),
            "\n".join(body_paras),
        )


def build_v15_extreme_ratio_fonts() -> None:
    """Test ROUND 11 PREDICTION at EXTREME usWinAscent ratios.

    Round 11 V14 covered 4 fonts in usWinAsc/upem range 0.8594–0.9951.
    V15 extends to extreme ratios discovered via probe_jp_font_ascent.py:
      Meiryo / Meiryo UI: usWinAsc=2171/2048 = 1.0601 → asc(14pt)=14.84pt
      Yu Gothic UI:        usWinAsc=2210/2048 = 1.0791 → asc(14pt)=15.11pt

    At base=14pt with raise=12pt, hps=7pt:
      Meiryo:        max(0, 12 + 5.25 - 14.841) = 2.409pt
      Yu Gothic UI:  max(0, 12 + 5.25 - 15.107) = 2.143pt
      (vs Yu Mincho 3.318pt, MS Mincho 5.219pt — clearly distinguishable
       at Word's 0.5pt rounding)

    Meiryo + Meiryo UI share TTF metrics (same TTC face); included as
    a sanity check that font name alone (not the underlying TTF) does
    not affect ruby ascent.
    """
    fonts = [
        ("Yu Gothic UI",  "YuGothicUI",  "Yu_GothicUI_1.0791"),
        ("メイリオ",        "Meiryo_jp",   "Meiryo_1.0601"),
        ("Meiryo",        "Meiryo_en",   "Meiryo_1.0601"),
        ("Meiryo UI",     "MeiryoUI",    "Meiryo_1.0601_(UI variant, same TTF)"),
    ]
    base_sz = 28
    base_pt = 14.0
    default_hps = base_sz // 2
    big_hps = base_sz
    for font_name, suffix, family in fonts:
        body_paras = [
            _para(_r(f"V15 P0 ref.", sz=str(base_sz), font_name=font_name)),
            _para(
                _r("P1 def: ", sz=str(base_sz), font_name=font_name),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(default_hps),
                    hps=str(default_hps),
                    font_name=font_name,
                ),
                _r("です。", sz=str(base_sz), font_name=font_name),
            ),
            _para(_r("V15 P2 ref.", sz=str(base_sz), font_name=font_name)),
            _para(
                _r("P3 r6: ", sz=str(base_sz), font_name=font_name),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(default_hps),
                    hps=str(default_hps),
                    hps_raise="12",
                    font_name=font_name,
                ),
                _r("です。", sz=str(base_sz), font_name=font_name),
            ),
            _para(_r("V15 P4 ref.", sz=str(base_sz), font_name=font_name)),
            _para(
                _r("P5 r12: ", sz=str(base_sz), font_name=font_name),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(default_hps),
                    hps=str(default_hps),
                    hps_raise="24",
                    font_name=font_name,
                ),
                _r("です。", sz=str(base_sz), font_name=font_name),
            ),
            _para(_r("V15 P6 ref.", sz=str(base_sz), font_name=font_name)),
            _para(
                _r("P7 hpsB: ", sz=str(base_sz), font_name=font_name),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(big_hps),
                    hps=str(big_hps),
                    font_name=font_name,
                ),
                _r("です。", sz=str(base_sz), font_name=font_name),
            ),
            _para(_r("V15 P8 close.", sz=str(base_sz), font_name=font_name)),
        ]
        write_docx(
            os.path.join(OUT_DIR, f"RUBY_V15_{suffix}_140dpt.docx"),
            "\n".join(body_paras),
        )


def build_v16_tier_bc_raise_sweep() -> None:
    """V15 (round 12) showed Round 11's usWinAscent formula breaks for
    tier B (Meiryo Regular) and tier C (Yu Gothic UI) fonts. V16 tests
    the formula STRUCTURE: does `expansion ∝ raise + 0.75×hps - constant`
    still hold linearly, just with a different constant per font?

    Focus: hold hps fixed at default (base/2 = 7pt), sweep raise across
    5 values to derive the raise→exp slope. If slope = 1.0 (per-pt raise
    yields per-pt exp), formula structure is preserved → tier B/C just
    needs per-font ascent calibration. If slope < 1.0, formula structure
    itself is wrong for these fonts.

    Hypothesis pre-V16:
    - V15 Yu Gothic UI: raise=6 → exp=0.5, raise=12 → exp=3.5
      slope estimate: (3.5 - 0.5) / (12 - 6) = 0.5 ← SUB-UNITY suggests
      structural difference

    Layout: per font, V13-triple chain
      P0 ref / P1 raise=6 / P2 ref / P3 raise=12 / P4 ref / P5 raise=18
      / P6 ref / P7 raise=24 / P8 ref / P9 raise=36 / P10 ref
    11 paragraphs per fixture, 2 fixtures.
    """
    fonts = [
        ("Yu Gothic UI", "YuGothicUI"),
        ("メイリオ",       "MeiryoRegular_jp"),
    ]
    raises_halfpt = [12, 24, 36, 48, 72]  # 6pt, 12pt, 18pt, 24pt, 36pt
    base_sz = 28
    base_pt = 14.0
    default_hps = base_sz // 2  # 14 halfpt = 7pt

    for font_name, suffix in fonts:
        body_paras = [_para(_r("V16 P0 ref.", sz=str(base_sz), font_name=font_name))]
        for i, rh in enumerate(raises_halfpt):
            r_pt = rh / 2
            body_paras.append(_para(
                _r(f"r{r_pt:g}: ", sz=str(base_sz), font_name=font_name),
                _ruby(
                    "含", "ふく",
                    base_sz=str(base_sz),
                    ruby_sz=str(default_hps),
                    hps=str(default_hps),
                    hps_raise=str(rh),
                    font_name=font_name,
                ),
                _r("です。", sz=str(base_sz), font_name=font_name),
            ))
            body_paras.append(_para(_r(f"V16 P{2*(i+1)} ref.", sz=str(base_sz), font_name=font_name)))
        write_docx(
            os.path.join(OUT_DIR, f"RUBY_V16_{suffix}_140dpt_raisesweep.docx"),
            "\n".join(body_paras),
        )


def build_v17_no_ruby_lineheight_per_font() -> None:
    """V17: pure no-ruby line-height profiling per font × base size.

    V14/V15 measurements showed Yu Mincho/Gothic at base=14pt has
    no_ruby_LH = 23.5pt (1.679× base) and Meiryo Regular = 27.375pt (1.955×
    base), violating the CLAUDE.md `base × 9/7 = 1.286× base` CJK rule
    derived from MS Mincho. V17 isolates no-ruby line-height by removing
    all ruby content, sweeping 5 base sizes per font.

    Per font: 11 paragraphs at 5 base sizes (2 paragraphs per size for
    clean dy + 1 closer). Pattern (per fixture):
      [P1 size_a / P2 size_a]   ← dy(P1,P2) = no_ruby_LH at a
      [P3 size_b / P4 size_b]   ← dy(P3,P4) = no_ruby_LH at b
      ... (5 bases × 2 = 10 paragraphs)
      [P11 closer]

    Note: dy(P2, P3) is base_a + spacing-with-base_b — discard. Only same-
    size pairs are valid.
    """
    fonts = [
        ("ＭＳ 明朝",       "MSMincho_control"),
        ("Yu Mincho",     "YuMincho"),
        ("游ゴシック",      "YuGothic"),
        ("Yu Gothic UI",  "YuGothicUI"),
        ("メイリオ",        "Meiryo"),
        ("Meiryo UI",     "MeiryoUI"),
    ]
    base_halfpts = [18, 21, 22, 24, 28]  # 9pt, 10.5pt, 11pt, 12pt, 14pt
    for font_name, suffix in fonts:
        body_paras = []
        for bhp in base_halfpts:
            bp = bhp / 2
            body_paras.append(_para(_r(f"V17 {bp}pt P_a 本文.", sz=str(bhp), font_name=font_name)))
            body_paras.append(_para(_r(f"V17 {bp}pt P_b 本文.", sz=str(bhp), font_name=font_name)))
        body_paras.append(_para(_r("V17 closer.", sz="21", font_name=font_name)))
        write_docx(
            os.path.join(OUT_DIR, f"RUBY_V17_{suffix}_no_ruby_LH.docx"),
            "\n".join(body_paras),
        )


def build_v18_no_docgrid_lineheight() -> None:
    """V18: pure no_ruby_LH WITHOUT <w:docGrid> element in sectPr.

    Round 15 audit revealed Oxi has a `lm0_lineauto.json` lookup table
    for "no docGrid" scenarios with values divergent from V17 (which
    uses docGrid type='default' linePitch=312):

      MS Mincho 14pt: LM0 lookup=21.0pt vs V17 measured=18.5pt (Δ +2.5)
      Yu Mincho 14pt: LM0=27.5pt vs V17=23.5pt (Δ +4.0)

    Hypothesis: LM0 represents Word's behavior at sectPr.docGrid ELEMENT
    ABSENT (vs V17 which has the element with type=default). Per spec
    §1.4, 'docGrid element absent → grid snap NOT applied' — so the
    line-height formula path may differ.

    V18 builds same V17-pattern (6 fonts × 5 base sizes) but with sectPr
    that omits <w:docGrid> entirely. If V18 == LM0 lookup table, LM0
    is validated for the no-docGrid scenario. If V18 == V17, LM0 may
    be incorrect or measure something else entirely.
    """
    fonts = [
        ("ＭＳ 明朝",       "MSMincho_control"),
        ("Yu Mincho",     "YuMincho"),
        ("游ゴシック",      "YuGothic"),
        ("Yu Gothic UI",  "YuGothicUI"),
        ("メイリオ",        "Meiryo"),
        ("Meiryo UI",     "MeiryoUI"),
    ]
    base_halfpts = [18, 21, 22, 24, 28]
    sect_pr_no_docgrid = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:cols w:space="425"/>'
        '</w:sectPr>'
    )
    for font_name, suffix in fonts:
        body_paras = []
        for bhp in base_halfpts:
            bp = bhp / 2
            body_paras.append(_para(_r(f"V18 {bp}pt P_a 本文.", sz=str(bhp), font_name=font_name)))
            body_paras.append(_para(_r(f"V18 {bp}pt P_b 本文.", sz=str(bhp), font_name=font_name)))
        body_paras.append(_para(_r("V18 closer.", sz="21", font_name=font_name)))
        body_xml = DOC_HEAD + "\n".join(body_paras) + sect_pr_no_docgrid + "\n</w:body>\n</w:document>"
        path = os.path.join(OUT_DIR, f"RUBY_V18_{suffix}_no_docgrid.docx")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CONTENT_TYPES)
            z.writestr("_rels/.rels", ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", DOC_RELS)
            z.writestr("word/document.xml", body_xml)
            z.writestr("word/styles.xml", STYLES_XML)
            z.writestr("word/settings.xml", SETTINGS_XML)
        print(f"  wrote {path}")


def build_v19_baseline_pattern_docgrid() -> None:
    """V19: replicate the EXACT baseline docGrid pattern.

    Survey (round 16) showed 49/51 baseline docs use:
      <w:docGrid w:linePitch="360"/>  (type attribute ABSENT)

    Oxi parser (parser/ooxml.rs:5532-5538) routes type-absent docGrid to
    `grid_line_pitch = None` (only "lines"/"linesAndChars" set Some).
    Combined with auto lineSpacing, this triggers `is_single_lm0`
    (layout/mod.rs:3294) → LM0 lookup table fires.

    LM0 vs V18 (no docGrid) divergence is 2.5–9pt. If V19 (baseline
    pattern) == V18 (no docGrid) measurements, LM0 is incorrect for
    nearly the entire baseline → significant SHIP OPPORTUNITY.

    V19 same as V17/V18 layout (6 fonts × 5 base sizes × 11 paragraphs)
    but with baseline-style sectPr.
    """
    fonts = [
        ("ＭＳ 明朝",       "MSMincho_control"),
        ("Yu Mincho",     "YuMincho"),
        ("游ゴシック",      "YuGothic"),
        ("Yu Gothic UI",  "YuGothicUI"),
        ("メイリオ",        "Meiryo"),
        ("Meiryo UI",     "MeiryoUI"),
    ]
    base_halfpts = [18, 21, 22, 24, 28]
    sect_pr_baseline = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:cols w:space="425"/>'
        '<w:docGrid w:linePitch="360"/>'  # exact baseline: type absent, linePitch=18pt
        '</w:sectPr>'
    )
    for font_name, suffix in fonts:
        body_paras = []
        for bhp in base_halfpts:
            bp = bhp / 2
            body_paras.append(_para(_r(f"V19 {bp}pt P_a 本文.", sz=str(bhp), font_name=font_name)))
            body_paras.append(_para(_r(f"V19 {bp}pt P_b 本文.", sz=str(bhp), font_name=font_name)))
        body_paras.append(_para(_r("V19 closer.", sz="21", font_name=font_name)))
        body_xml = DOC_HEAD + "\n".join(body_paras) + sect_pr_baseline + "\n</w:body>\n</w:document>"
        path = os.path.join(OUT_DIR, f"RUBY_V19_{suffix}_baseline_docgrid.docx")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CONTENT_TYPES)
            z.writestr("_rels/.rels", ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", DOC_RELS)
            z.writestr("word/document.xml", body_xml)
            z.writestr("word/styles.xml", STYLES_XML)
            z.writestr("word/settings.xml", SETTINGS_XML)
        print(f"  wrote {path}")


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
    build_v13_base_raise_grid()
    build_v14_font_family_variants()
    build_v15_extreme_ratio_fonts()
    build_v16_tier_bc_raise_sweep()
    build_v17_no_ruby_lineheight_per_font()
    build_v18_no_docgrid_lineheight()
    build_v19_baseline_pattern_docgrid()
    print("Done.")


if __name__ == "__main__":
    main()
