"""Bracket-pair compression isolation.

LW_20 showed 「...」 pairs compress. PI_C showed single 「 does NOT compress.
Test whether the pair pattern is the trigger.

Variants (Meiryo 10.5pt, useFE on, kern 3):
  BP_A: '観測値定義'×10 — control, no brackets (50 chars)
  BP_B: '「観測」'×8 + '観測' (単独) — '「観測」'={「 観 測 」}=4 chars × 8 = 32 + '観測'=2 = 34 chars
  BP_C: '観「測」'×10 — mid-line bracket pairs (40 chars)
  BP_D: '「観測値定」'×8 — enclosing brackets (48 chars), 「 at 0/6/12/... (mid-line except 0)
  BP_E: '観測「値」定'×8 — 「」 deeply mid-line (48 chars)
  BP_F: Halfwidth paren pair: '観(測)'×16 — () pairs
  BP_G: Fullwidth paren pair: '観（測）'×10 — （） pairs (40 chars)
  BP_H: Mixed punct: '観、測。値「定」'×6 (42 chars) — all punct types
"""
import os, zipfile

OUT_DIR = os.path.abspath("tools/metrics/bracket_pair_repro")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
SETTINGS = '''<?xml version="1.0"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>'''


def doc_xml(text):
    rpr = '<w:rFonts w:ascii="メイリオ" w:eastAsia="メイリオ" w:hAnsi="メイリオ"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:kern w:val="3"/>'
    esc = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return f'''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc}</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>'''


def build(label, text):
    path = os.path.join(OUT_DIR, f'{label}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', doc_xml(text))
    print(f'Built {label} text_len={len(text)}')


cases = [
    ('BP_A', '観測値定義' * 10),  # 50, no punct
    ('BP_B', '「観測」' * 8 + '観測'),  # 34 chars, 8 bracket-pairs + tail
    ('BP_C', '観「測」' * 10),  # 40 chars, 10 bracket-pairs mid-line
    ('BP_D', '「観測値定」' * 8),  # 48 chars, brackets enclose
    ('BP_E', '観測「値」定' * 8),  # 48 chars, brackets deeply mid-line
    ('BP_F', '観(測)' * 16),  # 64 chars — halfwidth parens mixed with CJK
    ('BP_G', '観（測）' * 10),  # 40 chars — fullwidth paren pairs
    ('BP_H', '観、測。値「定」' * 6),  # 42 chars — mixed punct
]
for (label, text) in cases:
    build(label, text)
