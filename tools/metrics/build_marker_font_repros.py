"""Build minimal repros for the body-marker font_family fix.

Scenario: a numbered list paragraph with a halfwidth marker (like "(1)")
where the paragraph's body text is in a CJK font. Word renders the marker
in the paragraph's font; Oxi previously used font_family=None, falling
back to a Latin default that produced narrower parens than Word.

Variants:
  MF_A: MS Mincho paragraph + halfwidth "(1)" marker
  MF_B: MS Gothic paragraph + halfwidth "(1)" marker
  MF_C: Meiryo paragraph + halfwidth "(1)" marker
  MF_D: MS Mincho paragraph + fullwidth "（１）" marker (sanity — should
        work unchanged because fullwidth glyph widths are the same)
"""
import os
import zipfile

OUT_DIR = os.path.abspath("tools/metrics/marker_font_repro")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'


def num_xml(lvl_text):
    return f'''<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="{lvl_text}"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>'''


def doc_xml(rpr, texts):
    paras = ''
    for t in texts:
        paras += (f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
                  f'<w:rPr>{rpr}</w:rPr></w:pPr>'
                  f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r></w:p>')
    return f'<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{paras}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>'


def build(label, lvl_text, rpr, texts):
    path = os.path.join(OUT_DIR, f'{label}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/numbering.xml', num_xml(lvl_text))
        z.writestr('word/document.xml', doc_xml(rpr, texts))
    print(f'Built {path}')


RPR_MINCHO = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/>'
RPR_GOTHIC = '<w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="21"/>'
RPR_MEIRYO = '<w:rFonts w:ascii="メイリオ" w:eastAsia="メイリオ" w:hAnsi="メイリオ"/><w:sz w:val="21"/>'

TEXTS = ["公開するデータの設計", "データを登録するシステム"]

build('MF_A', '(%1)', RPR_MINCHO, TEXTS)
build('MF_B', '(%1)', RPR_GOTHIC, TEXTS)
build('MF_C', '(%1)', RPR_MEIRYO, TEXTS)
build('MF_D', '（%1）', RPR_MINCHO, TEXTS)
