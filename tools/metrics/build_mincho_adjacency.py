"""Test adjacency rule on MS Mincho (Meiryo rule might be font-specific).

Build 3 key adjacency patterns in MS Mincho 10.5pt:
  MC_A: 観、「測 (test Rule A: 、「 → 、 compresses)
  MC_B: 観。）測 (test Rule A: 。) → 。 compresses)
  MC_C: 観）。測 (test Rule A: )。 → ) compresses)
  MC_D: 観）、測 (test Rule A: )、 → ) compresses)
  MC_E: 観、、測 (test: 、、 → first 、 compresses)
  MC_F: 観「「測 (test Rule B: 「「 → second 「 compresses)
  MC_CTRL: 観、測測 (control: 、測 → no compression, 、 stays 10.5)

Pattern: each is (item × 10) = 40 chars.
"""
import os, zipfile

OUT_DIR = os.path.abspath("tools/metrics/mincho_adjacency_repro")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
SETTINGS = '''<?xml version="1.0"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>'''


def doc_xml(text, font):
    rpr = f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:kern w:val="3"/>'
    esc = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return f'''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc}</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>'''


def build(label, text, font='ＭＳ 明朝'):
    path = os.path.join(OUT_DIR, f'{label}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', doc_xml(text, font))


# MS Mincho variants
build('MC_A_mincho', '観、「測' * 10, 'ＭＳ 明朝')  # 、「 compression
build('MC_B_mincho', '観。）測' * 10, 'ＭＳ 明朝')
build('MC_C_mincho', '観）。測' * 10, 'ＭＳ 明朝')
build('MC_D_mincho', '観）、測' * 10, 'ＭＳ 明朝')
build('MC_E_mincho', '観、、測' * 10, 'ＭＳ 明朝')
build('MC_F_mincho', '観「「測' * 10, 'ＭＳ 明朝')
build('MC_CTRL_mincho', '観、測測' * 10, 'ＭＳ 明朝')

# MS Gothic variants (another common font)
build('MC_A_gothic', '観、「測' * 10, 'ＭＳ ゴシック')
build('MC_CTRL_gothic', '観、測測' * 10, 'ＭＳ ゴシック')

# MS PGothic variants
build('MC_A_pgothic', '観、「測' * 10, 'ＭＳ Ｐゴシック')

# Yu Gothic variants
build('MC_A_yugothic', '観、「測' * 10, '游ゴシック')

# Meiryo control (confirms our earlier data)
build('MC_A_meiryo', '観、「測' * 10, 'メイリオ')
build('MC_CTRL_meiryo', '観、測測' * 10, 'メイリオ')

print("Built MC_* variants (mincho/gothic/pgothic/yugothic/meiryo)")
