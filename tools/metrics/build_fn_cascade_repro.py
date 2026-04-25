"""Build minimal repro docx fixtures for the b837 footnote cascade pattern.

The Layer 1 fix (`mod.rs:2049+` re-commit fn refs on mid-para break) needs ≥3
docs + minimal repro per Path B. b837 is the only baseline doc with fn refs;
these repros provide the minimal cases.

Pattern under test:
- A paragraph spans pages internally (mid-para page break inside layout).
- The paragraph has fn refs distributed across the break: some on OLD page,
  some on NEW page.
- Pre-fix bug: NEW-page refs' bodies are dropped because
  `footnote_reserve_current = 0` after MID_BREAK and the paragraph's NEW-page
  refs are never re-committed.

Scenarios:
  FN_CASCADE_A: ONE huge paragraph (~80 lines of CJK) with 6 fn refs evenly
                spread → page break occurs mid-paragraph → refs land on both
                pages.
  FN_CASCADE_B: TWO large paragraphs each spanning page break → refs
                distributed across 3 pages.
  FN_CASCADE_C: Edge case — a hanging-indent paragraph with fn refs
                (validates Layer 2 `.max(0.0)` clamp doesn't affect fn flow).

Verify with COM (Word) to confirm:
  1. Where Word puts the page break inside the paragraph
  2. Which refs land on which page
  3. Whether fn bodies render correctly on each page

Output dir: tools/metrics/fn_cascade_repro/
"""
import os, zipfile

OUT_DIR = r"tools\metrics\fn_cascade_repro"
os.makedirs(OUT_DIR, exist_ok=True)

CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"""

# Match b837 settings: compatibilityMode=15 + compressPunctuation
SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

# MS Mincho 10.5pt default to match b837
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
<w:style w:type="paragraph" w:styleId="FootnoteText">
<w:name w:val="footnote text"/>
<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
</w:style>
<w:style w:type="character" w:styleId="FootnoteReference">
<w:name w:val="footnote reference"/>
<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
</w:style>
</w:styles>"""


def footnotes_xml(fn_texts):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for i, t in enumerate(fn_texts, start=1):
        parts.append(
            f'<w:footnote w:id="{i}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
            f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            f'<w:r><w:t xml:space="preserve"> {t}</w:t></w:r></w:p></w:footnote>'
        )
    parts.append("</w:footnotes>")
    return "".join(parts)


def split_para_with_fns(text_segments, fn_ids_per_segment, indent_first=None, indent_left=None):
    """Build a single <w:p> with text segments and fn refs interleaved.

    text_segments: list of CJK text strings to render in order.
    fn_ids_per_segment: list of lists; fn_ids_per_segment[i] is the list of
        fn ids to insert AFTER segments[i].
    indent_first: optional firstLine indent in twips (positive = regular,
        negative = hanging).
    indent_left: optional left indent in twips.
    """
    parts = ["<w:p>"]
    if indent_first is not None or indent_left is not None:
        attrs = []
        if indent_left is not None:
            attrs.append(f'w:left="{indent_left}"')
        if indent_first is not None:
            sign = "firstLine" if indent_first >= 0 else "hanging"
            val = abs(indent_first)
            attrs.append(f'w:{sign}="{val}"')
        parts.append(f'<w:pPr><w:ind {" ".join(attrs)}/></w:pPr>')
    assert len(text_segments) == len(fn_ids_per_segment), \
        "segments and fn_ids_per_segment must align"
    for seg_text, fn_ids in zip(text_segments, fn_ids_per_segment):
        if seg_text:
            parts.append(f'<w:r><w:t xml:space="preserve">{seg_text}</w:t></w:r>')
        for fid in fn_ids:
            parts.append(
                f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                f'<w:footnoteReference w:id="{fid}"/></w:r>'
            )
    parts.append("</w:p>")
    return "".join(parts)


def filler_para(n_chars=400):
    """A normal CJK paragraph that takes ~ N chars."""
    base = "本文の充填段落です、改行が発生する程度の長さを確保します。"
    full = (base * (n_chars // len(base) + 1))[:n_chars]
    return f'<w:p><w:r><w:t xml:space="preserve">{full}</w:t></w:r></w:p>'


def document_xml(body_paras):
    body_inner = "".join(body_paras)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body_inner}'
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="425"/>'
        '<w:docGrid w:type="linesAndChars" w:linePitch="360"/>'
        '</w:sectPr>'
        '</w:body></w:document>'
    )


def build_docx(name, body_paras, fn_texts):
    files = {
        "[Content_Types].xml": CONTENT_TYPES_XML,
        "_rels/.rels": ROOT_RELS_XML,
        "word/_rels/document.xml.rels": DOC_RELS_XML,
        "word/styles.xml": STYLES_XML,
        "word/settings.xml": SETTINGS_XML,
        "word/footnotes.xml": footnotes_xml(fn_texts),
        "word/document.xml": document_xml(body_paras),
    }
    path = os.path.join(OUT_DIR, name)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for n, c in files.items():
            zf.writestr(n, c)
    print(f"Saved {path}")


def fn_cascade_a():
    """ONE huge paragraph (~3000 chars) with 6 fn refs spread evenly.

    Goal: paragraph body forces page break around the middle. Refs 1-2 land
    on page 1, refs 3-6 on page 2 (Word's actual distribution depends on
    where break occurs; verify with COM after build).

    Total ~3000 chars at MS Mincho 10.5pt → ~80 lines (40 chars/line) →
    ~3.5 pages of body. With ~50 lines/page, break at line 50 means
    chars ~2000 → refs at chars 500/1000/1500/2000/2500/2800 land roughly
    2 on page 1, 4 on page 2.
    """
    seg_chars = 500  # chars between fn refs
    n_refs = 6
    body_text = "footnote 前後の段落、各 fn は注を意味します、改行を確実に発生させるための充分な長さの段落を構成します。"
    # Build segments of ~seg_chars each
    segments = []
    for i in range(n_refs + 1):
        full = (body_text * (seg_chars // len(body_text) + 1))[:seg_chars]
        segments.append(full)
    fn_ids_per_segment = [[i + 1] if i < n_refs else [] for i in range(n_refs + 1)]
    para = split_para_with_fns(segments, fn_ids_per_segment)
    body = [para]
    fn_texts = [f"脚注 {i+1} の本文：簡潔な注釈テキストです。" for i in range(n_refs)]
    build_docx("FN_CASCADE_A.docx", body, fn_texts)


def fn_cascade_b():
    """TWO large paragraphs, each spanning a page break, refs across 3 pages.

    Para 1: ~1800 chars + 4 refs (1-4) → spans p1→p2.
    Filler para to position para 2.
    Para 2: ~1800 chars + 4 refs (5-8) → spans p2→p3.

    This tests TWO separate MID_BREAKs in a single doc, both with fn refs
    on the new pages.
    """
    body = []
    seg_chars = 450
    n_refs_per_para = 4
    body_text = "footnote 前後の段落、各 fn は注を意味します、改行を確実に発生させるための充分な長さの段落です。"

    # Para 1
    segments_1 = []
    for i in range(n_refs_per_para + 1):
        full = (body_text * (seg_chars // len(body_text) + 1))[:seg_chars]
        segments_1.append(full)
    fn_ids_1 = [[i + 1] if i < n_refs_per_para else [] for i in range(n_refs_per_para + 1)]
    body.append(split_para_with_fns(segments_1, fn_ids_1))

    # Filler — small para to nudge break point
    body.append(filler_para(80))

    # Para 2 — refs 5-8
    segments_2 = []
    for i in range(n_refs_per_para + 1):
        full = (body_text * (seg_chars // len(body_text) + 1))[:seg_chars]
        segments_2.append(full)
    fn_ids_2 = [[i + 5] if i < n_refs_per_para else [] for i in range(n_refs_per_para + 1)]
    body.append(split_para_with_fns(segments_2, fn_ids_2))

    fn_texts = [f"脚注 {i+1} の本文：注釈 #{i+1}。" for i in range(8)]
    build_docx("FN_CASCADE_B.docx", body, fn_texts)


def fn_cascade_c():
    """Edge case: hanging-indent paragraph with fn refs.

    Tests that Layer 2 `.max(0.0)` clamp on charGrid effective_first_indent
    does NOT affect fn-cascade behavior (Layer 1 fix is orthogonal). Also
    tests that hanging indent + fn refs renders correctly.
    """
    seg_chars = 500
    n_refs = 4
    body_text = "footnote ぶら下げインデント段落、各 fn は注を意味します。"
    segments = []
    for i in range(n_refs + 1):
        full = (body_text * (seg_chars // len(body_text) + 1))[:seg_chars]
        segments.append(full)
    fn_ids_per_segment = [[i + 1] if i < n_refs else [] for i in range(n_refs + 1)]
    # Hanging indent: left=480, firstLine=-240 (pulls first line back 240tw=12pt)
    para = split_para_with_fns(
        segments, fn_ids_per_segment,
        indent_left=480, indent_first=-240,
    )
    body = [para]
    fn_texts = [f"脚注 {i+1} 注釈本文。" for i in range(n_refs)]
    build_docx("FN_CASCADE_C.docx", body, fn_texts)


if __name__ == "__main__":
    fn_cascade_a()
    fn_cascade_b()
    fn_cascade_c()
    print("Done.")
