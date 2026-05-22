"""S181: build 8-variant minimal-repro matrix to map Word's row-height
border-allocation behavior across (num_tables, num_rows, trHeight,
insideH, tcMar) axes.

Goal: identify when Word allocates 0.5pt for outer-top or inside-H
border in row_height. S180 found `row_idx == 0` gate is too broad —
6 tokumei-family gains but 11 other docs regressed -0.0069 net mean
IoU. The Word behavior is doc-dependent; this matrix reverse-engineers
the rule.

Output: tools/golden-test/repros/row_border_matrix/RBM_<label>.docx

Variants:
  RBM_A  1 table  × 1 row,  border, no trHeight,    no tcMar    (= tokumei V101 unit)
  RBM_B  5 tables × 1 row,  border, no trHeight,    no tcMar    (= tokumei V101 pattern)
  RBM_C  5 tables × 1 row,  border, trHeight=660,   no tcMar    (= d4d126 row 1 pattern)
  RBM_D  1 table  × 5 rows, border + insideH, no trHeight, no tcMar (= multi-row, inside-H)
  RBM_E  1 table  × 5 rows, border NO insideH, no trHeight, no tcMar (= multi-row, no inside-H)
  RBM_F  1 table  × 5 rows, border + insideH, trHeight=660 on all, no tcMar (= multi-row + trHeight)
  RBM_G  5 tables × 1 row,  NO border,                              (control: 0 drift expected)
  RBM_H  5 tables × 1 row,  border, no trHeight, explicit tcMar top=0 (explicit override)

Each row: 1 cell, 1 paragraph with sb=87/sa=87 line=240 lineRule=exact.
Page A4, top/bottom margin 72pt, MS Mincho 10.5pt.

Run:
  python tools/metrics/gen_row_border_matrix.py
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'row_border_matrix')
os.makedirs(OUT_DIR, exist_ok=True)

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''
RELS_ROOT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="a">
    <w:name w:val="Normal"/>
    <w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
  </w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="compressPunctuation"/>
</w:settings>'''

PARA = ('<w:p>'
        '<w:pPr>'
        '<w:spacing w:beforeLines="30" w:before="87" w:afterLines="30" w:after="87" w:line="240" w:lineRule="exact"/>'
        '<w:rPr><w:spacing w:val="0"/></w:rPr>'
        '</w:pPr>'
        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
        '<w:t>R{r}</w:t></w:r>'
        '</w:p>')


def make_tbl_borders(has_border: bool, has_inside_h: bool, outer_sz: int = 4) -> str:
    if not has_border:
        return ''
    inner = (f'<w:top w:val="single" w:sz="{outer_sz}"/>'
             f'<w:left w:val="single" w:sz="{outer_sz}"/>'
             f'<w:bottom w:val="single" w:sz="{outer_sz}"/>'
             f'<w:right w:val="single" w:sz="{outer_sz}"/>')
    if has_inside_h:
        inner += '<w:insideH w:val="single" w:sz="4"/>'
    return f'<w:tblBorders>{inner}</w:tblBorders>'


def make_tcMar(explicit_zero: bool) -> str:
    if not explicit_zero:
        return ''
    return '<w:tcMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tcMar>'


def make_row(r_idx: int, tr_height: int | None, has_border: bool,
             explicit_tcMar: bool, tc_top_sz: int = 4) -> str:
    trpr = ''
    if tr_height is not None:
        trpr = f'<w:trPr><w:trHeight w:val="{tr_height}"/></w:trPr>'

    tcPr_parts = ['<w:tcW w:w="9000" w:type="dxa"/>']
    if has_border:
        tcPr_parts.append(
            f'<w:tcBorders>'
            f'<w:top w:val="single" w:sz="{tc_top_sz}"/>'
            f'<w:left w:val="single" w:sz="4"/>'
            f'<w:bottom w:val="single" w:sz="4"/>'
            f'<w:right w:val="single" w:sz="4"/>'
            f'</w:tcBorders>'
        )
    tcPr_parts.append(make_tcMar(explicit_tcMar))
    tcPr = '<w:tcPr>' + ''.join(p for p in tcPr_parts if p) + '</w:tcPr>'

    para = PARA.replace('{r}', str(r_idx))
    return f'<w:tr>{trpr}<w:tc>{tcPr}{para}</w:tc></w:tr>'


def make_table(num_rows: int, has_border: bool, has_inside_h: bool,
               tr_height: int | None, explicit_tcMar: bool,
               t_idx: int = 0, outer_sz: int = 4,
               first_row_top_sz: int | None = None) -> str:
    """first_row_top_sz: if set, overrides row 0's tc top border (matches
    7ead52's pattern where only row 0 has sz=12 top, rest use default)."""
    tbl_pr_parts = ['<w:tblW w:w="9000" w:type="dxa"/>']
    tbl_pr_parts.append(make_tbl_borders(has_border, has_inside_h, outer_sz))
    tbl_pr_parts.append('<w:tblLook w:val="04A0"/>')
    tbl_pr = '<w:tblPr>' + ''.join(p for p in tbl_pr_parts if p) + '</w:tblPr>'
    grid = '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
    rows = ''
    for i in range(num_rows):
        tc_top = first_row_top_sz if (i == 0 and first_row_top_sz is not None) else 4
        rows += make_row(t_idx * 10 + i, tr_height, has_border, explicit_tcMar, tc_top)
    return f'<w:tbl>{tbl_pr}{grid}{rows}</w:tbl>'


def make_doc(num_tables: int, num_rows_per_table: int,
             has_border: bool, has_inside_h: bool,
             tr_height: int | None, explicit_tcMar: bool,
             outer_sz: int = 4, first_row_top_sz: int | None = None) -> bytes:
    tables = ''.join(
        make_table(num_rows_per_table, has_border, has_inside_h, tr_height,
                   explicit_tcMar, t, outer_sz, first_row_top_sz)
        for t in range(num_tables)
    )
    body = f'{tables}<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>'
    sect_pr = ('<w:sectPr>'
               '<w:pgSz w:w="11906" w:h="16838"/>'
               '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>'
               '<w:docGrid w:type="lines" w:linePitch="360"/>'
               '</w:sectPr>')
    doc_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
               f'<w:body>{body}{sect_pr}</w:body>'
               '</w:document>')
    return doc_xml.encode('utf-8')


VARIANTS = [
    # (label, num_tables, num_rows, has_border, has_inside_h, tr_height, explicit_tcMar,
    #  outer_sz, first_row_top_sz)
    ('A_1tbl_1row',              1, 1, True,  False, None, False, 4,  None),
    ('B_5tbl_1row',              5, 1, True,  False, None, False, 4,  None),
    ('C_5tbl_1row_trH660',       5, 1, True,  False, 660,  False, 4,  None),
    ('D_1tbl_5row_insideH',      1, 5, True,  True,  None, False, 4,  None),
    ('E_1tbl_5row_noInsideH',    1, 5, True,  False, None, False, 4,  None),
    ('F_1tbl_5row_insH_trH',     1, 5, True,  True,  660,  False, 4,  None),
    ('G_5tbl_1row_NoBorder',     5, 1, False, False, None, False, 4,  None),
    ('H_5tbl_1row_tcMar0',       5, 1, True,  False, None, True,  4,  None),
    # S183: 7ead52-style variants (sz=12 outer border + trH + insideH)
    ('I_1tbl_5row_sz12_insH',    1, 5, True,  True,  None, False, 12, None),
    ('J_1tbl_5row_sz12_insH_trH', 1, 5, True, True,  860,  False, 12, None),
    ('K_1tbl_5row_sz12_noInsH_trH', 1, 5, True, False, 860, False, 12, None),
    # 7ead52's actual: sz=12 only on row 0 top (not on inside-H rows)
    ('L_1tbl_5row_row0sz12_trH', 1, 5, True,  True,  860,  False, 4,  12),
]


def write_docx(label: str, num_tables: int, num_rows: int,
               has_border: bool, has_inside_h: bool,
               tr_height: int | None, explicit_tcMar: bool,
               outer_sz: int = 4, first_row_top_sz: int | None = None):
    doc_bytes = make_doc(num_tables, num_rows, has_border, has_inside_h,
                          tr_height, explicit_tcMar, outer_sz, first_row_top_sz)
    path = os.path.join(OUT_DIR, f'RBM_{label}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', RELS_ROOT)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', doc_bytes)
    print(f'  wrote {path}')


def main():
    print(f'Writing {len(VARIANTS)} variants to {OUT_DIR}')
    for v in VARIANTS:
        write_docx(*v)


if __name__ == '__main__':
    main()
