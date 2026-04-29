"""
R101 minimal-repro: e8caed wp34 line-height hypothesis

R97 hypothesised "Para 34 wrap decision in break_into_lines for hanging-indent
CJK paragraph". R101 disproved that -- the actual bug is wp34's "備考" paragraph
height: Word reports h=10.5pt, Oxi computes 13.0pt (+2.5pt).

The pPr is unusual:
  <w:pPr>
    <w:snapToGrid w:val="0"/>            ← grid snap OFF
    <w:spacing w:beforeLines="50" w:before="136"/>  ← no line= attr (auto)
    <w:ind w:left="520" w:hangingChars="260" w:hanging="520"/>
  </w:pPr>
  <w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t> </w:t></w:r>      ← ASCII sp sz=20 (10pt)
  <w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>備考</w:t></w:r>   ← CJK sz=16 (8pt)

Hypothesis A: Word ignores ASCII whitespace runs when computing line height
              → only the 8pt CJK run drives line_height = 10.5pt (= 8 × 83/64
                CJK adjustment, ceil-to-0.5pt).
Hypothesis B: ASCII space contributes but at a different (smaller) effective
              size in Word's line-box logic.
Hypothesis C: Something about snapToGrid val=0 changes the run-merge rule.

This script builds 5 docx variants and measures each with Word COM.
"""
import os
import sys
import zipfile
import io
import time

# Force UTF-8 stdout (cp932 default on Windows-jp can't encode em-dash etc.)
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX_DIR = os.path.abspath(os.path.join(
    os.path.dirname(__file__), '..', 'fixtures', 'wp34_line_height_repro'))

# Minimal docx skeleton -- single section, MS Mincho default east-asia, A4 page
# with docGrid linePitch=272 (matches e8caed). Each variant differs only in
# the test paragraph's pPr/runs.
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

# styles.xml -- set ＭＳ 明朝 as default eastAsia, Century as default ascii
STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</w:styles>'''


def build_doc_xml(test_para_xml: str) -> str:
    # 3 paragraphs:
    # P1 = "ANCHOR1" (8pt CJK) -- top anchor
    # P2 = the test paragraph (the variant under test)
    # P3 = "ANCHOR2" (8pt CJK) -- bottom anchor; stride P2→P3 measures P2 height
    # All with snapToGrid val=0 for consistency.
    anchor_pre = (
        '<w:p><w:pPr>'
        '<w:snapToGrid w:val="0"/>'
        '</w:pPr>'
        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="16"/></w:rPr>'
        '<w:t>ア</w:t></w:r></w:p>'  # ア
    )
    anchor_post = (
        '<w:p><w:pPr>'
        '<w:snapToGrid w:val="0"/>'
        '</w:pPr>'
        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="16"/></w:rPr>'
        '<w:t>イ</w:t></w:r></w:p>'  # イ
    )
    sect_pr = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="567" w:right="1133" w:bottom="142" w:left="1133" w:header="0" w:footer="0" w:gutter="0"/>'
        '<w:cols w:space="425"/>'
        '<w:docGrid w:type="linesAndChars" w:linePitch="272"/>'
        '</w:sectPr>'
    )
    body = anchor_pre + test_para_xml + anchor_post + sect_pr
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {NS}>'
            f'<w:body>{body}</w:body>'
            '</w:document>')


def make_docx(path: str, test_para_xml: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES_XML)
        z.writestr('word/document.xml', build_doc_xml(test_para_xml))


# ------------------------------------------------------------------
# Variants
# ------------------------------------------------------------------
# Each test_para_xml is just the <w:p>...</w:p> for the test paragraph.

# v1 -- exact wp34 config: ASCII sp sz=20 + CJK "備考" sz=16, snapToGrid=0,
# no line= attr.
V1 = (
    '<w:p><w:pPr>'
    '<w:snapToGrid w:val="0"/>'
    '</w:pPr>'
    '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
    '<w:t xml:space="preserve"> </w:t></w:r>'
    '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="16"/></w:rPr>'
    '<w:t>備考</w:t></w:r>'  # 備考
    '</w:p>'
)

# v2 -- drop the leading ASCII space; only the CJK sz=16 run.
V2 = (
    '<w:p><w:pPr>'
    '<w:snapToGrid w:val="0"/>'
    '</w:pPr>'
    '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="16"/></w:rPr>'
    '<w:t>備考</w:t></w:r>'
    '</w:p>'
)

# v3 -- leading run is also CJK (eastAsia hint) at sz=16, same size as 備考.
V3 = (
    '<w:p><w:pPr>'
    '<w:snapToGrid w:val="0"/>'
    '</w:pPr>'
    '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="16"/></w:rPr>'
    '<w:t xml:space="preserve">　</w:t></w:r>'  # 　 fullwidth space
    '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="16"/></w:rPr>'
    '<w:t>備考</w:t></w:r>'
    '</w:p>'
)

# v4 -- both runs at sz=20 (10pt), bigger uniform size. Tests "do mixed-size
# rules disappear when sizes match".
V4 = (
    '<w:p><w:pPr>'
    '<w:snapToGrid w:val="0"/>'
    '</w:pPr>'
    '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
    '<w:t xml:space="preserve"> </w:t></w:r>'
    '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="20"/></w:rPr>'
    '<w:t>備考</w:t></w:r>'
    '</w:p>'
)

# v5 -- ASCII space sz=16 (matches CJK size). Tests "does ASCII contribute
# when same size as CJK".
V5 = (
    '<w:p><w:pPr>'
    '<w:snapToGrid w:val="0"/>'
    '</w:pPr>'
    '<w:r><w:rPr><w:sz w:val="16"/></w:rPr>'
    '<w:t xml:space="preserve"> </w:t></w:r>'
    '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="16"/></w:rPr>'
    '<w:t>備考</w:t></w:r>'
    '</w:p>'
)

# v6 -- control: ASCII run is non-space (a letter "X") at sz=20.
# Tests "is ASCII space treated specially or is ASCII letter at sz=20 also
# 'ignored'?".
V6 = (
    '<w:p><w:pPr>'
    '<w:snapToGrid w:val="0"/>'
    '</w:pPr>'
    '<w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
    '<w:t xml:space="preserve">X </w:t></w:r>'
    '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="16"/></w:rPr>'
    '<w:t>備考</w:t></w:r>'
    '</w:p>'
)

VARIANTS = [
    ('v1_ascii_sp20_cjk16', V1, 'wp34 exact config: ASCII sp sz=20 + CJK 備考 sz=16'),
    ('v2_cjk16_only', V2, 'CJK 備考 sz=16 only (no leading space)'),
    ('v3_cjk_fullwidth_sp16_cjk16', V3, 'CJK fullwidth sp sz=16 + CJK 備考 sz=16'),
    ('v4_ascii_sp20_cjk20', V4, 'ASCII sp sz=20 + CJK 備考 sz=20 (both 10pt)'),
    ('v5_ascii_sp16_cjk16', V5, 'ASCII sp sz=16 + CJK 備考 sz=16 (matched 8pt)'),
    ('v6_ascii_X20_cjk16', V6, 'ASCII letter X sz=20 + CJK 備考 sz=16'),
]


def measure_via_com(docx_paths):
    """Open each in Word; return per-doc {p1_y, p2_y, p3_y, stride_p2}."""
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    results = {}
    try:
        for name, path in docx_paths:
            print(f"\n=== {name} ===")
            doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            try:
                # Use collapsed-start range per R30 fix.
                ys = []
                for i in range(1, 4):
                    para = doc.Paragraphs(i)
                    rng = para.Range
                    rng_start = doc.Range(rng.Start, rng.Start)
                    y = rng_start.Information(6)  # wdVerticalPositionRelativeToPage
                    ys.append(y)
                    text = rng.Text[:30].replace('\r', '\\r')
                    print(f"  P{i} y={y:7.3f}pt  '{text}'")
                stride_p2 = ys[2] - ys[1]
                stride_p1 = ys[1] - ys[0]
                print(f"  stride P1→P2 = {stride_p1:.3f}pt (anchor row)")
                print(f"  stride P2→P3 = {stride_p2:.3f}pt (TEST = effective P2 height)")
                results[name] = {
                    'p1_y': ys[0], 'p2_y': ys[1], 'p3_y': ys[2],
                    'stride_p1_to_p2': stride_p1,
                    'stride_p2_to_p3': stride_p2,  # effective height of test paragraph
                }
            finally:
                doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    return results


def main():
    print("Building variants...")
    paths = []
    for name, body, desc in VARIANTS:
        path = os.path.join(DOCX_DIR, f'{name}.docx')
        make_docx(path, body)
        paths.append((name, path))
        print(f"  built: {name}.docx -- {desc}")

    print("\nMeasuring via Word COM...")
    results = measure_via_com(paths)

    print("\n=== Summary ===")
    print(f"{'variant':<35} stride_p2_to_p3 (= effective height of test paragraph)")
    for name, _, desc in VARIANTS:
        r = results.get(name, {})
        s = r.get('stride_p2_to_p3', float('nan'))
        print(f"  {name:<35} {s:7.3f}pt  ← {desc}")

    # Save raw to JSON for archiving.
    import json
    out = os.path.join(DOCX_DIR, '_measurement_results.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nResults saved to {out}")


if __name__ == '__main__':
    main()
