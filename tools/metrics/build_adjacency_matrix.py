"""Systematic adjacency compression test.

For each (prev, next) pair of punctuation/brackets, measure whether
the NEXT char compresses. Build pattern: CJK prev next CJK prev next CJK...

Test pairs (ordered prev→next):
  Rows = prev char, Columns = next char
  Using: 、 。 「 」 （ ） ． ，

Build repro: '観' + prev + next + '測' repeated 10 times = 40 chars per repro.
Measures width of NEXT char in Word.

Naming: ADJ_<prev><next> e.g., ADJ_COM_LBK = 、「
Using ASCII labels for filenames.
"""
import os, zipfile

OUT_DIR = os.path.abspath("tools/metrics/adjacency_matrix_repro")
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


PUNCTS = {
    'CM':  '、',   # U+3001 comma
    'PD':  '。',   # U+3002 period
    'LBK': '「',   # U+300C left bracket
    'RBK': '」',   # U+300D right bracket
    'LPN': '（',   # U+FF08 fullwidth left paren
    'RPN': '）',   # U+FF09 fullwidth right paren
    'FPD': '．',   # U+FF0E fullwidth period
    'FCM': '，',   # U+FF0C fullwidth comma
}

# Generate all pairs (prev, next)
pairs = []
for prev_label, prev_ch in PUNCTS.items():
    for next_label, next_ch in PUNCTS.items():
        if prev_label == next_label and prev_ch in '、。．，':
            # Skip nonsense pairs like 、、 or 。。
            continue
        label = f'ADJ_{prev_label}_{next_label}'
        # Pattern: 観 prev next 測 — 4 chars cycle
        text = ('観' + prev_ch + next_ch + '測') * 10  # 40 chars
        pairs.append((label, text, prev_ch, next_ch))
        build(label, text)

print(f'Built {len(pairs)} adjacency repros in {OUT_DIR}')
