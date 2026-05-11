"""Day 33 part 36 — Factor C minimal repro generator.

Hypothesis (from analyze_factor_c_row_trajectory output):
  For table cells with v_align=center + lh_rule=exact + fs=10.5,
  Oxi's row height is ~2.35pt SHORTER than Word's per-row, accumulating
  to ~-14pt over 6 rows (observed consistently in d4d126/6514/a1d6 t1).

Repro variants isolate which attribute drives the -2.35pt/row delta:
  v01: 10 rows × 1 cell, v_align=center, lh=exact line=240 (12pt), fs=10.5
       — matches d4d126/6514/a1d6 t1 setup. Expected: -2.35pt/row.
  v02: control — v_align=top, lh=exact line=240, fs=10.5
       — isolates v_align contribution.
  v03: control — v_align=center, lh=auto (no exact), fs=10.5
       — isolates lh_rule contribution.
  v04: control — v_align=center, lh=exact line=240, fs=12
       — isolates fs contribution.
  v05: control — v_align=center, lh=exact line=300 (15pt), fs=10.5
       — isolates lh_val contribution.

Output: tools/golden-test/repros/factor_c/v0N.docx for N in 1..5.
Measurement: tools/metrics/measure_factor_c_repro.py.
"""
from __future__ import annotations
import os, sys, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

OUT = Path('tools/golden-test/repros/factor_c')
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
<w:sz w:val="21"/><w:szCs w:val="24"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style>
</w:styles>'''


def cell_text(label):
    """Single-line text in a paragraph that wraps in a cell paragraph properties.

    pPr inside the cell uses lineRule (set at the variant level via param)."""
    return f'<w:t>{label}</w:t>'


def make_row(label, lh_rule, lh_val, fs_half, v_align, cell_width_tw=8000):
    """Build a single-row, single-cell row of a table.

    lh_rule: 'auto' or 'exact'
    lh_val: line value in twips (e.g. 240 = 12pt)
    fs_half: half-point size (21 = 10.5pt, 24 = 12pt)
    v_align: 'top' / 'center' / 'bottom'
    """
    lh_rule_attr = 'auto' if lh_rule == 'auto' else 'exact'
    pPr_spacing = (f'<w:spacing w:line="{lh_val}" w:lineRule="{lh_rule_attr}"/>'
                   if lh_rule != 'auto' else '')
    p_xml = (f'<w:p><w:pPr>{pPr_spacing}</w:pPr>'
             f'<w:r><w:rPr><w:sz w:val="{fs_half}"/></w:rPr>'
             f'<w:t>{label}</w:t></w:r></w:p>')
    cell_xml = (f'<w:tc>'
                f'<w:tcPr><w:tcW w:w="{cell_width_tw}" w:type="dxa"/>'
                f'<w:vAlign w:val="{v_align}"/></w:tcPr>'
                f'{p_xml}'
                f'</w:tc>')
    row_xml = f'<w:tr>{cell_xml}</w:tr>'
    return row_xml


def make_table(rows_xml):
    return ('<w:tbl>'
            '<w:tblPr>'
            '<w:tblW w:w="8000" w:type="dxa"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '</w:tblBorders>'
            '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
            f'{rows_xml}'
            '</w:tbl>')


def make_document(variant_rows_xml):
    table_xml = make_table(variant_rows_xml)
    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397" w:gutter="0"/>'
            '<w:docGrid w:type="lines" w:linePitch="360"/>'
            '</w:sectPr>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {NS}><w:body>{table_xml}<w:p/>{sect}</w:body></w:document>')


def build_docx(name, n_rows, lh_rule, lh_val, fs_half, v_align):
    p = OUT / f'{name}.docx'
    if p.exists(): p.unlink()
    rows = ''.join(make_row(f'row{i+1:02d}', lh_rule, lh_val, fs_half, v_align)
                   for i in range(n_rows))
    doc_xml = make_document(rows)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/document.xml', doc_xml)
    print(f'  wrote {p}')


if __name__ == '__main__':
    print('Factor C minimal repro generator — 5 variants:')
    print('  v01: v_align=center, lh=exact line=240 (12pt), fs=10.5 — TARGET')
    build_docx('v01_center_exact240_fs10p5', n_rows=10,
               lh_rule='exact', lh_val=240, fs_half=21, v_align='center')

    print('  v02 (control): v_align=top, lh=exact line=240, fs=10.5')
    build_docx('v02_top_exact240_fs10p5', n_rows=10,
               lh_rule='exact', lh_val=240, fs_half=21, v_align='top')

    print('  v03 (control): v_align=center, lh=auto, fs=10.5')
    build_docx('v03_center_auto_fs10p5', n_rows=10,
               lh_rule='auto', lh_val=0, fs_half=21, v_align='center')

    print('  v04 (control): v_align=center, lh=exact line=240, fs=12')
    build_docx('v04_center_exact240_fs12', n_rows=10,
               lh_rule='exact', lh_val=240, fs_half=24, v_align='center')

    print('  v05 (control): v_align=center, lh=exact line=300 (15pt), fs=10.5')
    build_docx('v05_center_exact300_fs10p5', n_rows=10,
               lh_rule='exact', lh_val=300, fs_half=21, v_align='center')

    print('done.')
