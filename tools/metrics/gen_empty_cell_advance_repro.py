"""Day 33 part 51 — Empty paragraph in cell advance minimal repros (R7.3).

Hypothesis: Oxi over-pumps empty paragraphs in cells by ~1-2pt vs Word,
which when accumulated over de6e's 69 empty paragraphs = 50-100pt of
"offsetting bug" Day 33 part 17 was hiding.

Each repro: 2-row table.
  row 1: text paragraph (anchor)
  row 2: empty paragraph (vary attributes)

We measure: advance from row 1 first-char y to row 2 first-char y.

Variants:
  e01: empty row 2 no explicit attrs (uses default fs from docDefault=10pt)
  e02: empty row 2 with explicit fs=10pt + lineRule=auto
  e03: empty row 2 with explicit fs=10.5 + lineRule=exact 12pt
  e04: empty row 2 with sb=87 + lineRule=exact 12pt
  e05: empty row 2 with sb=87 sa=87 lineRule=exact 12pt (备考-like)
  e06: empty row 2 with fs=10.5 lh=auto (no exact, sb=87)
  e07: same as e05 but auto trHeight (no trh)
  e08: same as e05 but exact trHeight=658
"""
from __future__ import annotations
import os, sys, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

OUT = Path('tools/golden-test/repros/empty_cell_advance')
OUT.mkdir(parents=True, exist_ok=True)

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
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
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>
<w:sz w:val="21"/><w:szCs w:val="21"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>
</w:styles>'''


def make_text_row(label, trHeight=None):
    """Text row with `label` content."""
    trPr = f'<w:trPr><w:trHeight w:val="{trHeight}"/></w:trPr>' if trHeight else ''
    pPr = '<w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr>'
    p_xml = f'<w:p>{pPr}<w:r><w:t>{label}</w:t></w:r></w:p>'
    tcPr = ('<w:tcPr><w:tcW w:w="8000" w:type="dxa"/>'
            '<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>'
            '<w:vAlign w:val="center"/></w:tcPr>')
    return f'<w:tr>{trPr}<w:tc>{tcPr}{p_xml}</w:tc></w:tr>'


def make_empty_row(*, pPr_extra='', rPr='', trHeight=None):
    """Empty paragraph row."""
    trPr = f'<w:trPr><w:trHeight w:val="{trHeight}"/></w:trPr>' if trHeight else ''
    pPr = f'<w:pPr>{pPr_extra}</w:pPr>'
    p_xml = f'<w:p>{pPr}{rPr}</w:p>'  # empty paragraph (no run)
    tcPr = ('<w:tcPr><w:tcW w:w="8000" w:type="dxa"/>'
            '<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>'
            '<w:vAlign w:val="center"/></w:tcPr>')
    return f'<w:tr>{trPr}<w:tc>{tcPr}{p_xml}</w:tc></w:tr>'


def make_table(rows):
    return ('<w:tbl>'
            '<w:tblPr>'
            '<w:tblW w:w="8000" w:type="dxa"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '</w:tblBorders>'
            '<w:tblLayout w:type="fixed"/>'
            '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
            + ''.join(rows)
            + '</w:tbl>')


def make_doc(table_xml):
    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/>'
            '</w:sectPr>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {NS}><w:body>{table_xml}<w:p/>{sect}</w:body></w:document>')


def build_docx(name, rows_xml):
    p = OUT / f'{name}.docx'
    if p.exists(): p.unlink()
    table_xml = make_table(rows_xml)
    doc_xml = make_doc(table_xml)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/document.xml', doc_xml)
    print(f'  wrote {p}')


if __name__ == '__main__':
    print('Empty cell paragraph advance repros:')

    # Each: row 1 = text "anchor", row 2 = empty paragraph
    print('  e01: empty row 2 no attrs')
    build_docx('e01_empty_default',
               [make_text_row('row1'),
                make_empty_row()])

    print('  e02: empty row 2 sz=20 (10pt) lineRule=auto')
    build_docx('e02_empty_fs10_auto',
               [make_text_row('row1'),
                make_empty_row(rPr='<w:rPr><w:sz w:val="20"/></w:rPr>')])

    print('  e03: empty row 2 sz=21 (10.5pt) lineRule=exact 12pt')
    build_docx('e03_empty_fs10p5_exact240',
               [make_text_row('row1'),
                make_empty_row(pPr_extra='<w:spacing w:line="240" w:lineRule="exact"/>',
                               rPr='<w:rPr><w:sz w:val="21"/></w:rPr>')])

    print('  e04: empty row 2 sb=87 lineRule=exact 12pt sz=21')
    build_docx('e04_empty_sb_only',
               [make_text_row('row1'),
                make_empty_row(pPr_extra='<w:spacing w:before="87" w:line="240" w:lineRule="exact"/>',
                               rPr='<w:rPr><w:sz w:val="21"/></w:rPr>')])

    print('  e05: empty row 2 sb=87 sa=87 lineRule=exact 12pt sz=21 (备考-like)')
    build_docx('e05_empty_sb_sa',
               [make_text_row('row1'),
                make_empty_row(pPr_extra='<w:spacing w:before="87" w:after="87" w:line="240" w:lineRule="exact"/>',
                               rPr='<w:rPr><w:sz w:val="21"/></w:rPr>')])

    print('  e06: empty row 2 sb=87 sa=87 lineRule=auto sz=21')
    build_docx('e06_empty_sb_sa_auto_lh',
               [make_text_row('row1'),
                make_empty_row(pPr_extra='<w:spacing w:before="87" w:after="87"/>',
                               rPr='<w:rPr><w:sz w:val="21"/></w:rPr>')])

    print('  e07: 备考-like empty + auto trHeight (no trh)')
    build_docx('e07_empty_sb_sa_auto_trh',
               [make_text_row('row1'),
                make_empty_row(pPr_extra='<w:spacing w:before="87" w:after="87" w:line="240" w:lineRule="exact"/>',
                               rPr='<w:rPr><w:sz w:val="21"/></w:rPr>',
                               trHeight=None)])

    print('  e08: same as e05 with trHeight=658 (atLeast)')
    build_docx('e08_empty_sb_sa_trh658',
               [make_text_row('row1'),
                make_empty_row(pPr_extra='<w:spacing w:before="87" w:after="87" w:line="240" w:lineRule="exact"/>',
                               rPr='<w:rPr><w:sz w:val="21"/></w:rPr>',
                               trHeight=658)])

    print('done.')
