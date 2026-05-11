"""Day 33 part 38 — Row-count sweep to nail exact Word spacing.before rule.

Day 33 part 37 reproduced -2.46pt/row drift in v06 (7-row d4d126 mirror)
and identified spacing.before as the discriminator. Also re-measured
Day 33 part 17's R1A_spacing_lineRule.docx and found:

  R1A_baseline       (no spacing)         r1 y=43.0
  R1A_spacing_before (only spacing.before) r1 y=46.5  -> +3.5pt added
  R1A_spacing_lineRule (before + lh=exact) r1 y=46.5  -> +3.5pt added
  R1A_lineRule_exact   (only lh=exact)     r1 y=43.0  -> 0pt added

Word ADDS +3.5pt for spacing.before=87 in 1-row tables. Day 33 part 17's
claim "Word always suppresses spacing.before" is contradicted by direct
re-measurement.

This campaign tests: does Word's behavior depend on row count?
  v14: 1 row  (same pPr as v06: before=87 + lh=240 exact + vAlign=center + trHeight=658)
  v15: 2 rows
  v16: 3 rows
  v17: 5 rows
  v18: same as v06 (7 rows; control)

Measurement: per-row word_y; identify if any row has anomalous advance.

Hypothesis A: Word applies +3.5pt for ALL rows in ALL tables (no row-count dep).
  -> Day 33 part 17's fix is fundamentally wrong; correct fix = apply +3.5pt
  not 0pt for all cell paragraphs.

Hypothesis B: Word's behavior differs by row count.
  -> Need to identify exact rule.
"""
from __future__ import annotations
import sys, os, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

OUT = Path('tools/golden-test/repros/factor_c')

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
<w:style w:type="paragraph" w:customStyle="1" w:styleId="ac"><w:name w:val="一太郎"/>
<w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/>
<w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/>
<w:spacing w:line="210" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="ＭＳ 明朝"/>
<w:spacing w:val="-1"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>
</w:styles>'''


def make_row(label_idx):
    """Single-cell row, same pPr as v06 (target d4d126 conditions)."""
    tcPr = ('<w:tcPr><w:tcW w:w="9343" w:type="dxa"/>'
            '<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>'
            '<w:vAlign w:val="center"/></w:tcPr>')
    pPr = ('<w:pPr><w:pStyle w:val="ac"/>'
           '<w:spacing w:beforeLines="30" w:before="87" w:afterLines="30" w:after="87" '
           'w:line="240" w:lineRule="exact"/></w:pPr>')
    p_xml = (f'<w:p>{pPr}<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
             f'<w:t>row{label_idx:02d}</w:t></w:r></w:p>')
    return f'<w:tr><w:tc>{tcPr}{p_xml}</w:tc></w:tr>'


def make_table(n_rows):
    return ('<w:tbl>'
            '<w:tblPr>'
            '<w:tblW w:w="9343" w:type="dxa"/>'
            '<w:tblInd w:w="433" w:type="dxa"/>'
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
            '<w:tblGrid><w:gridCol w:w="9343"/></w:tblGrid>'
            + ''.join(make_row(i + 1) for i in range(n_rows))
            + '</w:tbl>')


def make_doc(n_rows):
    table = make_table(n_rows)
    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/>'
            '</w:sectPr>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {NS}><w:body>{table}<w:p/>{sect}</w:body></w:document>')


def build_docx(name, n_rows):
    p = OUT / f'{name}.docx'
    if p.exists(): p.unlink()
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/document.xml', make_doc(n_rows))
    print(f'  wrote {p}')


if __name__ == '__main__':
    print('Row-count sweep (same pPr as v06):')
    build_docx('v14_1row', 1)
    build_docx('v15_2row', 2)
    build_docx('v16_3row', 3)
    build_docx('v17_5row', 5)
    # v18 same as v06 = 7 rows; skip (already built)
    print('done.')
