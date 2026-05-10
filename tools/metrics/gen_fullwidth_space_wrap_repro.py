"""Day 33 part 19 — Test whether Word wraps all-whitespace paragraphs.

Hypothesis: Word does NOT wrap paragraphs containing only fullwidth spaces
(or only whitespace). Oxi does, causing +54pt drift in d77a58 wi=129
(142 fullwidth spaces rendered as 4 lines vs Word's 1 line).

Variants:
  WS_10:  10  fullwidth spaces (within line width)
  WS_50:  50  fullwidth spaces (1 line at fs=12 / 487pt body width = 600pt > body)
  WS_100: 100 fullwidth spaces (clearly exceeds body width)
  WS_142: 142 fullwidth spaces (matches d77a58 wi=129)
  WS_300: 300 fullwidth spaces (very long)
  MIX_10: 10 fullwidth spaces followed by Japanese text (control)
"""
from __future__ import annotations
import os, zipfile
from pathlib import Path

OUT = Path('tools/golden-test/repros/fullwidth_space_wrap')
OUT.mkdir(parents=True, exist_ok=True)

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
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

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:eastAsia="MS Gothic"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>'''


def make_para_with_spaces(n_spaces, suffix=''):
    spaces = '　' * n_spaces
    text = spaces + suffix
    return ('<w:p><w:pPr>'
            '<w:ind w:firstLineChars="100" w:firstLine="240"/>'
            '</w:pPr>'
            '<w:r><w:rPr><w:sz w:val="24"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def make_document(test_para):
    pre = '<w:p><w:r><w:t>BEFORE marker</w:t></w:r></w:p>'
    post = '<w:p><w:r><w:t>AFTER marker</w:t></w:r></w:p>'
    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:docGrid w:type="lines" w:linePitch="312"/>'
            '</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{pre}{test_para}{post}{sect}</w:body></w:document>'


def build_docx(name, test_para):
    p = OUT / f'{name}.docx'
    if p.exists(): p.unlink()
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/document.xml', make_document(test_para))
    print(f'  wrote {p}')


if __name__ == '__main__':
    build_docx('WS_10', make_para_with_spaces(10))
    build_docx('WS_50', make_para_with_spaces(50))
    build_docx('WS_100', make_para_with_spaces(100))
    build_docx('WS_142', make_para_with_spaces(142))
    build_docx('WS_300', make_para_with_spaces(300))
    build_docx('MIX_10_TEXT', make_para_with_spaces(10, 'テキスト終わり'))
    build_docx('MIX_50_TEXT', make_para_with_spaces(50, 'テキスト終わり'))
