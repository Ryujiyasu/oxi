"""Layer 2 sub-layer investigation (Session N+2): derive Word's actual
charGrid + indent line-1 wrap budget formula via controlled minimal repros.

Hypothesis 2a (from `docs/spec/b837_fn_cascade_multi_session_plan.md`):
  Word's line-1 budget = (charsPerLine - leftChars - firstLineChars) * pitch

Oxi's current formula:
  Layer 2 (1-line patch) tried `current_width starts at first_line_indent`
  but burasagari hangs the trailing char so wrap doesn't fire.
  Layer 2a tighter formula: replace `(total_cells - (indent_cells - 1)) * pitch`
  with `(charsPerLine - leftChars - firstLineChars) * pitch`.

This file builds a 10-repro matrix to pin down Word's actual rule:

Series A — vary leftChars (firstLineChars=0):
  L2A_0: leftChars=0
  L2A_1: leftChars=1
  L2A_2: leftChars=2 (b837 PARA 49 left)
  L2A_3: leftChars=3

Series B — vary firstLineChars (leftChars=0):
  L2B_1: firstLineChars=1 (b837 PARA 49 firstLine)
  L2B_2: firstLineChars=2

Series C — combined (mimics b837 PARA 49):
  L2C_20: leftChars=2, firstLineChars=0
  L2C_21: leftChars=2, firstLineChars=1 (= b837 PARA 49 indent shape)
  L2C_22: leftChars=2, firstLineChars=2

Series D — chars vs twips dominance (covers session 38c question):
  L2D_inconsist: leftChars=2 + left=600tw (inconsistent: chars→24pt, twips→30pt)

All paragraphs use:
  - 50 fullwidth CJK chars (long enough to definitely wrap)
  - MS Mincho 10.5pt (matches b837 default)
  - Same docGrid linesAndChars linePitch=360 + 1418tw margins
  - compatibilityMode=15 + characterSpacingControl=compressPunctuation
  - Final char "。" (tests burasagari trigger)

After building, run COM via `measure_chargrid_indent_repro.py` to get
the char N where Word breaks line 1 for each repro.

Output dir: tools/metrics/chargrid_indent_repro/
"""
import os, zipfile

OUT_DIR = r"tools\metrics\chargrid_indent_repro"
os.makedirs(OUT_DIR, exist_ok=True)

CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

# Match b837 settings for fair comparison
SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

# MS Mincho 10.5pt default (sz val="21" = 10.5pt) to match b837
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="ＭＳ 明朝" w:cs="Times New Roman"/>
<w:sz w:val="21"/><w:szCs w:val="21"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr>
<w:spacing w:line="276" w:lineRule="auto"/>
</w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

# 50 fullwidth CJK ideographs ending with "。" — 49 chars + "。"
# Use a distinct, well-known sequence so per-char tracing is easy
BODY_CHARS = "本文の充填段落です改行が発生する程度の長さを確保しますこれは検証用の段落でし。"
# Pad to 50 chars
while len(BODY_CHARS) < 50:
    BODY_CHARS = BODY_CHARS[:-1] + "の" + BODY_CHARS[-1]
BODY_CHARS = BODY_CHARS[:50]
assert len(BODY_CHARS) == 50, f"got {len(BODY_CHARS)} chars"
assert BODY_CHARS[-1] == "。", f"last char is {BODY_CHARS[-1]!r}"

# Sanity: fix manually if the splice trick fails
BODY_CHARS = "本文の充填段落です改行が発生する程度の長さを確保することで折り返し位置を精密に測定します。"
BODY_CHARS = BODY_CHARS[:50]
if len(BODY_CHARS) < 50:
    BODY_CHARS = (BODY_CHARS[:-1] + "の" * (50 - len(BODY_CHARS)) + BODY_CHARS[-1])
assert len(BODY_CHARS) == 50, f"got {len(BODY_CHARS)} chars"


def make_indent(left_twips=None, first_twips=None, left_chars=None, first_chars=None):
    """Build a <w:ind> attribute string."""
    attrs = []
    if left_chars is not None:
        attrs.append(f'w:leftChars="{left_chars}"')
    if left_twips is not None:
        attrs.append(f'w:left="{left_twips}"')
    if first_chars is not None:
        # firstLineChars OR hangingChars depending on sign — for our positive
        # values we always use firstLineChars
        attrs.append(f'w:firstLineChars="{first_chars}"')
    if first_twips is not None and first_twips >= 0:
        attrs.append(f'w:firstLine="{first_twips}"')
    elif first_twips is not None and first_twips < 0:
        attrs.append(f'w:hanging="{abs(first_twips)}"')
    if not attrs: return ""
    return f'<w:ind {" ".join(attrs)}/>'


def make_para(text, ind_xml):
    pPr = f'<w:pPr>{ind_xml}</w:pPr>' if ind_xml else ""
    return (
        f'<w:p>{pPr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


def document_xml(body_paras):
    body_inner = "".join(body_paras)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body_inner}'
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1418" w:bottom="1440" w:left="1418" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="425"/>'
        '<w:docGrid w:type="linesAndChars" w:linePitch="360"/>'
        '</w:sectPr>'
        '</w:body></w:document>'
    )


def build_docx(name, ind_kwargs):
    para = make_para(BODY_CHARS, make_indent(**ind_kwargs))
    body = [para]
    files = {
        "[Content_Types].xml": CONTENT_TYPES_XML,
        "_rels/.rels": ROOT_RELS_XML,
        "word/_rels/document.xml.rels": DOC_RELS_XML,
        "word/styles.xml": STYLES_XML,
        "word/settings.xml": SETTINGS_XML,
        "word/document.xml": document_xml(body),
    }
    path = os.path.join(OUT_DIR, name)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for n, c in files.items():
            zf.writestr(n, c)
    print(f"Saved {path}  (ind: {ind_kwargs})")


# ---------------- Series A: vary leftChars (no firstLine) ----------------
build_docx("L2A_0.docx", {})
build_docx("L2A_1.docx", {"left_chars": 100, "left_twips": 240})  # 1 char ≈ 12pt
build_docx("L2A_2.docx", {"left_chars": 200, "left_twips": 480})  # 2 chars ≈ 24pt
build_docx("L2A_3.docx", {"left_chars": 300, "left_twips": 720})  # 3 chars ≈ 36pt

# ---------------- Series B: vary firstLineChars (no left) ----------------
build_docx("L2B_1.docx", {"first_chars": 100, "first_twips": 240})
build_docx("L2B_2.docx", {"first_chars": 200, "first_twips": 480})

# ---------------- Series C: combined ----------------
build_docx("L2C_20.docx", {"left_chars": 200, "left_twips": 480})  # = L2A_2 (alias)
build_docx("L2C_21.docx", {"left_chars": 200, "left_twips": 480, "first_chars": 100, "first_twips": 240})  # b837 PARA 49 shape
build_docx("L2C_22.docx", {"left_chars": 200, "left_twips": 480, "first_chars": 200, "first_twips": 480})

# ---------------- Series D: chars vs twips inconsistency ----------------
build_docx("L2D_inconsist.docx", {"left_chars": 200, "left_twips": 600})  # chars→24pt, twips→30pt

print("\nDone. Body text is:")
print(f"  {BODY_CHARS!r} ({len(BODY_CHARS)} chars)")
print("\nRun measure_chargrid_indent_repro.py next to capture Word's per-repro break point.")
