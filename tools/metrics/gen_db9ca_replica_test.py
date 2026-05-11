"""Day 33 part 30 (2026-05-11) — Targeted minimal repro replicating
db9ca wi=37's exact properties to find the overflow-tolerance differentiator.

db9ca wi=37 properties:
- pStyle custom "31" basedOn Normal "a"
- Normal has widowControl="0"
- ind: leftChars=202 left=426 hanging=2
- Multi-run paragraph with Times New Roman + MS Pゴシック rFonts
- Latin English text, 221 chars wrapping to 3 lines
- Contains <w:lastRenderedPageBreak/> inside runs (Word save-time marker)

Variants test each property in isolation.
"""
from __future__ import annotations
import os, zipfile
from pathlib import Path

OUT = Path('tools/golden-test/repros/db9ca_replica')
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


def make_styles_db9ca_like():
    """Replicate db9ca styles: Normal "a" with widowControl=0, custom "31" with indent."""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="31"><w:name w:val="表 (緑)  31"/>
<w:basedOn w:val="a"/><w:pPr><w:ind w:leftChars="400" w:left="840"/></w:pPr></w:style>
</w:styles>'''


def make_fill_para(idx):
    return f'<w:p><w:r><w:t>F{idx:02d}</w:t></w:r></w:p>'


# Variants - different test paragraphs

def test_para_tnr_simple(text):
    """Simple TNR run."""
    return (f'<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" '
            f'w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>')


def test_para_tnr_indented(text):
    """TNR + db9ca-like hanging indent."""
    return (f'<w:p><w:pPr>'
            f'<w:ind w:leftChars="202" w:left="426" w:hanging="2"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" '
            f'w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r></w:p>')


def test_para_style31(text):
    """Use style "31" (db9ca's actual style)."""
    return (f'<w:p><w:pPr><w:pStyle w:val="31"/>'
            f'<w:ind w:leftChars="202" w:left="426" w:hanging="2"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" '
            f'w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/>'
            f'<w:color w:val="000000"/></w:rPr><w:t>{text}</w:t></w:r></w:p>')


def test_para_style31_multirun(text):
    """Style "31" + multi-run with mixed rPr (like db9ca wi=37)."""
    n = len(text)
    third = n // 3
    return (f'<w:p><w:pPr><w:pStyle w:val="31"/>'
            f'<w:ind w:leftChars="202" w:left="426" w:hanging="2"/></w:pPr>'
            # Run 1: black
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" '
            f'w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/>'
            f'<w:color w:val="000000"/></w:rPr><w:t xml:space="preserve">{text[:third]}</w:t></w:r>'
            # Run 2: underline
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" '
            f'w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/>'
            f'<w:color w:val="000000"/><w:u w:val="single"/></w:rPr>'
            f'<w:t xml:space="preserve">{text[third:2*third]}</w:t></w:r>'
            # Run 3: plain black again
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" '
            f'w:eastAsia="ＭＳ Ｐゴシック" w:hAnsi="Times New Roman"/>'
            f'<w:color w:val="000000"/></w:rPr><w:t xml:space="preserve">{text[2*third:]}</w:t></w:r>'
            f'</w:p>')


def make_document(fills, test_para_xml):
    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397" w:gutter="0"/>'  # db9ca margins
            '<w:docGrid w:type="lines" w:linePitch="360"/>'
            '</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{fills}{test_para_xml}{sect}</w:body></w:document>'


def build_docx(name, n_fill, test_xml):
    p = OUT / f'{name}.docx'
    if p.exists(): p.unlink()
    fills = ''.join(make_fill_para(i) for i in range(n_fill))
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', make_styles_db9ca_like())
        z.writestr('word/document.xml', make_document(fills, test_xml))
    print(f'  wrote {p}')


# Long Latin text similar to db9ca wi=37 length & content
LONG_LATIN = (
    '"This Terms of Use" herein does not apply to the following Content. '
    'If there is any Content that clearly states that different terms of use apply, '
    'it is described in "Important note on Public Data License (Version 1.0)".'
)


if __name__ == '__main__':
    # db9ca margin: top=1418tw=70.9pt, bottom=1418tw=70.9pt
    # Content height ~700pt. 700/18 ≈ 38 lines.
    # With 70.9pt top margin, line 1 at y=70.9, line 38 at y=70.9+37*18=736.9.
    # Test at fill=38: test_para at y=736.9+18=754.9 page 1.
    # Visible bottom 754.9+10.5=765.4 ≤ 771.1 page_bottom → should fit visibly.
    # Grid advance 754.9+18=772.9 > 771.1 → fails strict check.

    for n_fill in [37, 38, 39]:
        build_docx(f'DBR_tnr_simple_fill{n_fill}', n_fill, test_para_tnr_simple(LONG_LATIN))
        build_docx(f'DBR_tnr_indented_fill{n_fill}', n_fill, test_para_tnr_indented(LONG_LATIN))
        build_docx(f'DBR_style31_fill{n_fill}', n_fill, test_para_style31(LONG_LATIN))
        build_docx(f'DBR_style31_multirun_fill{n_fill}', n_fill, test_para_style31_multirun(LONG_LATIN))
    print(f'Wrote 12 repros to {OUT}')
