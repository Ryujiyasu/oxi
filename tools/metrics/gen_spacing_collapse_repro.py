"""Day 33 part 50 — Spacing collapse minimal repros (R7.2).

Test how Word handles spacing.before / spacing.after BETWEEN
consecutive table rows. Hypothesis: Word collapses (max not additive).

Variants:
  s01: 2 rows, sb=87 sa=0 lineRule=exact 12pt vAlign=center, trHeight=658(atLeast)
  s02: 2 rows, sb=0 sa=87 lineRule=exact 12pt (only after)
  s03: 2 rows, sb=87 sa=87 lineRule=exact 12pt (both)
  s04: 4 rows, sb=87 sa=87 (more rows to see cumulative)
  s05: 4 rows, sb=0 sa=0 (control)
  s06: 4 rows, sb=87 sa=87, NO trHeight (auto)
"""
from __future__ import annotations
import os, sys, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

OUT = Path('tools/golden-test/repros/spacing_collapse')
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
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style>
</w:styles>'''


def make_row(label, sb, sa, trHeight=None, vAlign='center'):
    """Make a single-cell row."""
    trPr = f'<w:trPr><w:trHeight w:val="{trHeight}"/></w:trPr>' if trHeight else ''
    spacing_attrs = []
    if sb > 0:
        spacing_attrs.append(f'w:before="{sb}"')
    if sa > 0:
        spacing_attrs.append(f'w:after="{sa}"')
    spacing_attrs.append('w:line="240" w:lineRule="exact"')
    spacing = f'<w:spacing {" ".join(spacing_attrs)}/>'
    pPr = f'<w:pPr>{spacing}</w:pPr>'
    p_xml = f'<w:p>{pPr}<w:r><w:t>{label}</w:t></w:r></w:p>'
    tcPr = (f'<w:tcPr><w:tcW w:w="8000" w:type="dxa"/>'
            f'<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>'
            f'<w:vAlign w:val="{vAlign}"/></w:tcPr>')
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


def build_docx(name, *, n_rows, sb, sa, trHeight=None, vAlign='center'):
    p = OUT / f'{name}.docx'
    if p.exists(): p.unlink()
    rows = [make_row(f'row{i+1}', sb, sa, trHeight=trHeight, vAlign=vAlign) for i in range(n_rows)]
    table_xml = make_table(rows)
    doc_xml = make_doc(table_xml)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/document.xml', doc_xml)
    print(f'  wrote {p}')


if __name__ == '__main__':
    print('Spacing collapse repros:')
    print('  s01: 2r sb=87 sa=0 lh=exact240 trh=658 (atLeast)')
    build_docx('s01_2r_sb_only_trh658', n_rows=2, sb=87, sa=0, trHeight=658)
    print('  s02: 2r sb=0 sa=87 lh=exact240 trh=658')
    build_docx('s02_2r_sa_only_trh658', n_rows=2, sb=0, sa=87, trHeight=658)
    print('  s03: 2r sb=87 sa=87 lh=exact240 trh=658')
    build_docx('s03_2r_sb_sa_trh658', n_rows=2, sb=87, sa=87, trHeight=658)
    print('  s04: 4r sb=87 sa=87 lh=exact240 trh=658')
    build_docx('s04_4r_sb_sa_trh658', n_rows=4, sb=87, sa=87, trHeight=658)
    print('  s05: 4r sb=0 sa=0 lh=exact240 trh=658 (control)')
    build_docx('s05_4r_nospacing_trh658', n_rows=4, sb=0, sa=0, trHeight=658)
    print('  s06: 4r sb=87 sa=87 lh=exact240 NO trh (auto)')
    build_docx('s06_4r_sb_sa_auto', n_rows=4, sb=87, sa=87, trHeight=None)
    print('done.')
