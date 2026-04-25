"""Per-punct isolation: measure Word's compression behavior for each CJK punctuation.

For each punctuation char P, build a doc with: (CJK × 4) P repeated to > 50 chars,
so we can observe the per-P width via per-char X diff.

Pattern: "観測値定P" × 10 = 50 chars (1 P every 5 chars, 10 P occurrences)

Punctuations tested:
  PI_A: 、 (U+3001, ideographic comma)
  PI_B: 。 (U+3002, ideographic period)
  PI_C: 「 (U+300C, left corner bracket)
  PI_D: 」 (U+300D, right corner bracket)
  PI_E: （ (U+FF08, fullwidth left paren)
  PI_F: ） (U+FF09, fullwidth right paren)
  PI_G: ． (U+FF0E, fullwidth full stop)
  PI_H: ， (U+FF0C, fullwidth comma)
  PI_I: ？ (U+FF1F, fullwidth question mark)
  PI_J: ： (U+FF1A, fullwidth colon)
  PI_K: ； (U+FF1B, fullwidth semicolon)
  PI_L: 〜 (U+301C, wave dash)
  PI_M: ・ (U+30FB, katakana middle dot)

Control:
  PI_0: no punct, pure CJK 50 chars (baseline)

All docs use useFELayout=on + kern=3 (matching e3c545 compat).
Single-paragraph, Meiryo 10.5pt.
"""
import os
import zipfile

OUT_DIR = os.path.abspath("tools/metrics/punct_isolation_repro")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'

SETTINGS = '''<?xml version="1.0"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat>
    <w:useFELayout/>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
  </w:compat>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>'''


def doc_xml(text: str) -> str:
    rpr = '<w:rFonts w:ascii="メイリオ" w:eastAsia="メイリオ" w:hAnsi="メイリオ"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:kern w:val="3"/>'
    esc = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return f'''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc}</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>'''


def build(label: str, text: str):
    path = os.path.join(OUT_DIR, f'{label}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', doc_xml(text))
    print(f'Built {path} text_len={len(text)}')


# Control: pure CJK, no punct
PI_0 = '観測値定義' * 10  # 50 chars, no punct

# Per-punct isolation: "観測値定P" × 10 = 50 chars with 10 P occurrences
def pattern(p):
    return ('観測値定' + p) * 10  # 5 chars × 10 = 50

cases = [
    ('PI_0', PI_0, 'no punct (control)'),
    ('PI_A', pattern('、'), 'U+3001 ideographic comma'),
    ('PI_B', pattern('。'), 'U+3002 ideographic period'),
    ('PI_C', pattern('「'), 'U+300C left corner bracket'),
    ('PI_D', pattern('」'), 'U+300D right corner bracket'),
    ('PI_E', pattern('（'), 'U+FF08 fullwidth left paren'),
    ('PI_F', pattern('）'), 'U+FF09 fullwidth right paren'),
    ('PI_G', pattern('．'), 'U+FF0E fullwidth full stop'),
    ('PI_H', pattern('，'), 'U+FF0C fullwidth comma'),
    ('PI_I', pattern('？'), 'U+FF1F fullwidth question mark'),
    ('PI_J', pattern('：'), 'U+FF1A fullwidth colon'),
    ('PI_K', pattern('；'), 'U+FF1B fullwidth semicolon'),
    ('PI_L', pattern('〜'), 'U+301C wave dash'),
    ('PI_M', pattern('・'), 'U+30FB katakana middle dot'),
]

for label, text, desc in cases:
    build(label, text)
    # print description for reference
    print(f'  {label}: {desc}')
