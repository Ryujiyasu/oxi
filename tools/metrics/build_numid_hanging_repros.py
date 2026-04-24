"""Build minimal repro docs for numId+hanging text-X spec.

Variants:
  NH_A: decimalFullWidth ('３．'), left=426 hanging=426 (21.3pt)
  NH_B: decimal ('3.'), left=720 hanging=360 (36pt/18pt)
  NH_C: decimal with suff=space, left=720 hanging=360
  NH_D: bullet ('●'), left=720 hanging=360
  NH_E: decimalFullWidth, left=567 hanging=567 (28.35pt)
  NH_F: decimal, left=720 hanging=720 (equal, CJK common)

Hypothesis: in all variants, Word places first text char at x=LeftIndent.
Marker width never pushes text position further.
"""
import os
import zipfile

OUT_DIR = os.path.abspath("tools/metrics/numid_hanging_repro")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'


def num_xml(num_fmt, lvl_text, left, hanging, suff=None):
    suff_tag = f'<w:suff w:val="{suff}"/>' if suff else ''
    return f'''<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="{num_fmt}"/>
      <w:lvlText w:val="{lvl_text}"/>
      <w:lvlJc w:val="left"/>
      {suff_tag}
      <w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>'''


def doc_xml(rpr, text, ind_override=None):
    ind = f'<w:ind {ind_override}/>' if ind_override else ''
    paras = ''
    for t in text:
        paras += (f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
                  f'{ind}<w:rPr>{rpr}</w:rPr></w:pPr>'
                  f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r></w:p>')
    return f'<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{paras}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>'


def build(label, num_fmt, lvl_text, left, hanging, rpr, text, suff=None, ind_override=None):
    path = os.path.join(OUT_DIR, f'{label}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/numbering.xml', num_xml(num_fmt, lvl_text, left, hanging, suff))
        z.writestr('word/document.xml', doc_xml(rpr, text, ind_override))
    print(f'Built {path}')


RPR_MINCHO = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/>'
RPR_MEIRYO = '<w:rFonts w:ascii="メイリオ" w:eastAsia="メイリオ" w:hAnsi="メイリオ"/><w:sz w:val="21"/>'

TEXT = ["基本的な考え方", "公開するデータ", "オントロジ"]

build('NH_A', 'decimalFullWidth', '%1．', 426, 426, RPR_MINCHO, TEXT)
build('NH_B', 'decimal', '%1.', 720, 360, RPR_MINCHO, TEXT)
build('NH_C', 'decimal', '%1.', 720, 360, RPR_MINCHO, TEXT, suff='space')
build('NH_D', 'bullet', '&#x25CF;', 720, 360, RPR_MINCHO, TEXT)
build('NH_E', 'decimalFullWidth', '%1．', 567, 567, RPR_MEIRYO, TEXT)
build('NH_F', 'decimal', '%1.', 720, 720, RPR_MINCHO, TEXT)
