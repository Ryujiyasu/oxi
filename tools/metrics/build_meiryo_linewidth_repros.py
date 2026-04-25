"""Build minimal repro docs to isolate Word's Meiryo line-level compression.

Goal: e3c545 P30 fits 47 chars in 454pt in Word but Oxi computes 470.82pt.
~16pt line-level compression is unexplained.

Hypotheses to test:
  H1: useFELayout (FE) triggers compression of fullwidth punctuation mid-line
  H2: kern (GPOS kern pairs) alone causes the compression
  H3: specific punct chars squeeze (「」 bracket pair, 、。 commas)
  H4: (1) Latin-parens group has special handling near CJK

Variants — each is a single-paragraph doc with controlled content. Measure
Word line width via COM per-char scan.

  LW_00: 50 CJK chars (no punct), Meiryo 10.5pt, useFELayout=off, kern=off
  LW_01: 50 CJK chars (no punct), Meiryo 10.5pt, useFELayout=on,  kern=off
  LW_02: 50 CJK chars (no punct), Meiryo 10.5pt, useFELayout=off, kern=3
  LW_03: 50 CJK chars (no punct), Meiryo 10.5pt, useFELayout=on,  kern=3

  LW_10: 40 CJK + 10 fullwidth punct mid-line, useFELayout=off, kern=off
  LW_11: 40 CJK + 10 fullwidth punct mid-line, useFELayout=on,  kern=off
  LW_12: 40 CJK + 10 fullwidth punct mid-line, useFELayout=off, kern=3
  LW_13: 40 CJK + 10 fullwidth punct mid-line, useFELayout=on,  kern=3

  LW_20: isolate 「」 bracket pair (〇「〇〇〇〇〇〇〇〇」×5, useFE=on kern=3)
  LW_21: isolate 、comma (〇〇、〇〇、〇〇、〇〇、〇〇、〇〇、〇〇、〇〇、〇〇、)
  LW_22: isolate 。period (〇〇〇〇。〇〇〇〇。〇〇〇〇。...)
  LW_23: isolate ． fullwidth period (same pattern)

  LW_30: mimic e3c545 P30 structure (with useFE=on kern=3, characterSpacingControl=doNotCompress)

If LW_00=LW_01=LW_02=LW_03 (all same width) → pure CJK never compresses (expected).
If LW_10=LW_11=LW_12=LW_13 → punctuation never compresses (then gap comes from elsewhere).
If LW_13 < LW_10 strictly → useFE+kern together compress punct.
If LW_11 < LW_10 → useFE alone compresses punct.
If LW_20, LW_21, LW_22 differ from their uncompressed expectation → per-char compression quantified.
"""
import os
import zipfile

OUT_DIR = os.path.abspath("tools/metrics/meiryo_linewidth_repro")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'


def settings_xml(use_fe_layout: bool, kern: int = 0):
    """Settings with optional useFELayout and kern threshold."""
    fe_tag = "<w:useFELayout/>" if use_fe_layout else ""
    # kern is set per-run, not in settings — we'll set it per-rPr
    return f'''<?xml version="1.0"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat>
    {fe_tag}
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
  </w:compat>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>'''


def doc_xml(text: str, kern_val: int = 0):
    """Single paragraph with given text, Meiryo 10.5pt, optional kern."""
    kern_tag = f'<w:kern w:val="{kern_val}"/>' if kern_val else ''
    rpr = f'<w:rFonts w:ascii="メイリオ" w:eastAsia="メイリオ" w:hAnsi="メイリオ"/><w:sz w:val="21"/><w:szCs w:val="21"/>{kern_tag}'
    # Escape text for XML
    esc = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return f'''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc}</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>'''


def build(label: str, text: str, use_fe_layout: bool, kern: int):
    path = os.path.join(OUT_DIR, f'{label}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/settings.xml', settings_xml(use_fe_layout, kern))
        z.writestr('word/document.xml', doc_xml(text, kern))
    print(f'Built {path} (useFE={use_fe_layout}, kern={kern}, text_len={len(text)})')


# Text sets
PURE_CJK_50 = '観測値観測値観測値観測値観測値観測値観測値観測値観測値観測値'  # 50 CJK, no punct
# 50 chars: 観測値×10 = 30, need 50. Use longer repetition
PURE_CJK_50 = '観測値のデータを定義します統計観測値のデータを定義します統計観測値のデ'  # ~48

# Actually let's make it exactly 50 CJK chars
PURE_CJK_50 = ('観測値のデータ統計' * 6)[:50]  # 観測値のデータ統計 = 9 chars, ×6 = 54, trim to 50
assert len(PURE_CJK_50) == 50, f'got {len(PURE_CJK_50)}'

# Text with fullwidth punctuation mid-line (40 CJK + 10 punct = 50 chars)
CJK_WITH_PUNCT_50 = 'データは、各機関で独自に定義します。具体例は、「観測値の定義」を参照してください'
# Let's verify — want 40 CJK + 10 fullwidth punct
# Count: デ1 ー2 タ3 は4 、5 各6 機7 関8 で9 独10 自11 に12 定13 義14 し15 ま16 す17 。18 具19 体20 例21 は22 、23 「24 観25 測26 値27 の28 定29 義30 」31 を32 参33 照34 し35 て36 く37 だ38 さ39 い40
# That's 40 chars and includes 5 punct (、。、「」) = only 5 punct. Need 10.
# Let me reconstruct with MORE punctuation:
CJK_WITH_PUNCT_50 = 'データ、値、例、項、図、表、式、値、数、定、の、は、に、を、が、で、と、も、や、、'
# 20 cycles of "X、" pattern = 40 chars = 20 X + 20 、, but that gives 40 chars not 50
# Let's just use the real e3c545 P30 text variant

# For isolation: LW_20/21/22 — pattern-controlled
LW_20_TEXT = '「観測値の統計定義」' * 5  # 10 chars × 5 = 50, with 「 」 pairs
LW_21_TEXT = '観測値、' * 10  # 4 chars × 10 = 40 chars, half are 、
LW_22_TEXT = '観測値定義。' * 8  # 6 × 8 = 48, with 。
LW_23_TEXT = '観測値定義．' * 8  # 6 × 8 = 48, with fullwidth period ．

# Series 0: pure CJK, no punct — should NOT compress
build('LW_00', PURE_CJK_50, False, 0)
build('LW_01', PURE_CJK_50, True,  0)
build('LW_02', PURE_CJK_50, False, 3)
build('LW_03', PURE_CJK_50, True,  3)

# Series 1: CJK with mid-line punct — 50 chars total
# Make 40 CJK + 10 punct inline
CJK_PUNCT = '観測値、各機関で独自に定義し、具体例は「観測値の定義」を参照し、以下の表を見ます。'
# Count: 観1 測2 値3 、4 各5 機6 関7 で8 独9 自10 に11 定12 義13 し14 、15 具16 体17 例18 は19 「20 観21 測22 値23 の24 定25 義26 」27 を28 参29 照30 し31 、32 以33 下34 の35 表36 を37 見38 ま39 す40 。41
# 41 chars total. Punctuation: 、、「」、。 = 6 punct. OK close enough.
build('LW_10', CJK_PUNCT, False, 0)
build('LW_11', CJK_PUNCT, True,  0)
build('LW_12', CJK_PUNCT, False, 3)
build('LW_13', CJK_PUNCT, True,  3)

# Series 2: punct isolation (useFE=on, kern=3 — the suspected trigger combo)
build('LW_20', LW_20_TEXT, True, 3)  # many 「」 pairs
build('LW_21', LW_21_TEXT, True, 3)  # many 、
build('LW_22', LW_22_TEXT, True, 3)  # many 。
build('LW_23', LW_23_TEXT, True, 3)  # many ． fullwidth period

# Series 3: e3c545 P30 exact mimic
E3C545_P30 = 'メタデータは、各機関で独自に定義します。具体例は、「９．例 (1)メタデータ」を参照ください。'
# 47 chars in the real doc
build('LW_30', E3C545_P30, True, 3)  # useFE + kern (matches e3c545 settings)
build('LW_31', E3C545_P30, False, 0)  # no useFE, no kern (control)
