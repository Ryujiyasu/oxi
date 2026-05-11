"""Day 33 part 37 — Faithful d4d126 t1 mirror + incremental simplifications.

Phase R1.6: Day 33 part 36 v01-v05 minimal repros didn't reproduce real-doc
trajectory. This script builds v06 (faithful mirror of d4d126 t1) then v07-v11
that incrementally strip features to identify the -2.35pt/row discriminator.

d4d126 t1 features (extracted from real docx):
- sectPr: pgMar top=1440 right=1080 bottom=1440 left=1080 footer=992
          docGrid type=linesAndChars linePitch=292 charSpace=1453
- styles: 'ac' (一太郎) basedOn nothing, widowControl=0 wordWrap=0
          autoSpaceDE/DN=0 adjustRightInd=0, lh=210 exact, fs=21 (10.5pt),
          MS Mincho. Also Normal 'a' with widowControl=0.
- tblPr: tblInd=433, tblLayout=fixed
- rows 2-7 have pPr: beforeLines=30 before=87 afterLines=30 after=87
         line=240 lineRule=exact
- row 1 has gridSpan=4, vAlign=center, bottom=dashed border
- mixed trHeight values: r2=658 r4=437 r6=549 (others auto)
- mostly vAlign=center

Variants:
  v06: full mirror (target — should reproduce -2.35pt/row)
  v07: v06 - style 'ac' (Normal 'a' for paragraphs)
  v08: v06 - docGrid (default linePitch=360 type=default)
  v09: v06 - vAlign=center (vAlign=top)
  v10: v06 - spacing.before/after (only line=240 lineRule=exact)
  v11: v06 - trHeight rules (all auto)
  v12: v06 - tblLayout=fixed (auto)
  v13: v06 - all body content before table (just bare table at body start)
"""
from __future__ import annotations
import os, sys, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

OUT = Path('tools/golden-test/repros/factor_c')
OUT.mkdir(parents=True, exist_ok=True)

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')

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

STYLES_FULL = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
<w:qFormat/><w:rsid w:val="006068C3"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="ac"><w:name w:val="一太郎"/>
<w:rsid w:val="00A54881"/>
<w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/>
<w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/>
<w:spacing w:line="210" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="ＭＳ 明朝"/>
<w:spacing w:val="-1"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>
</w:styles>'''


def make_pPr(label_idx, with_spacing=True, with_pstyle_ac=True, with_after=True):
    pstyle = '<w:pStyle w:val="ac"/>' if with_pstyle_ac else ''
    if with_spacing:
        if with_after:
            spacing = ('<w:spacing w:beforeLines="30" w:before="87" '
                       'w:afterLines="30" w:after="87" w:line="240" w:lineRule="exact"/>')
        else:
            spacing = ('<w:spacing w:beforeLines="30" w:before="87" '
                       'w:line="240" w:lineRule="exact"/>')
    else:
        spacing = '<w:spacing w:line="240" w:lineRule="exact"/>'
    return f'<w:pPr>{pstyle}{spacing}</w:pPr>'


def make_cell_para(label, pPr):
    return (f'<w:p>{pPr}<w:r>'
            f'<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            f'<w:t>{label}</w:t></w:r></w:p>')


def make_row(label_idx, trHeight, v_align, with_spacing=True, with_pstyle_ac=True,
             with_after=True, gridSpan=None):
    """One row with one cell.

    trHeight: int twips or None (auto).
    v_align: 'top' | 'center' | 'bottom' | None.
    gridSpan: int or None.
    """
    trPr = f'<w:trPr><w:trHeight w:val="{trHeight}"/></w:trPr>' if trHeight else ''
    gridSpan_xml = f'<w:gridSpan w:val="{gridSpan}"/>' if gridSpan else ''
    vAlign_xml = f'<w:vAlign w:val="{v_align}"/>' if v_align else ''
    tcPr = (f'<w:tcPr><w:tcW w:w="9343" w:type="dxa"/>{gridSpan_xml}'
            f'<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>'
            f'{vAlign_xml}</w:tcPr>')
    pPr = make_pPr(label_idx, with_spacing=with_spacing,
                   with_pstyle_ac=with_pstyle_ac, with_after=with_after)
    p_xml = make_cell_para(f'row{label_idx:02d}', pPr)
    cell_xml = f'<w:tc>{tcPr}{p_xml}</w:tc>'
    row_xml = f'<w:tr>{trPr}{cell_xml}</w:tr>'
    return row_xml


def make_table(rows_xml, tblLayout_fixed=True):
    tblLayout = '<w:tblLayout w:type="fixed"/>' if tblLayout_fixed else ''
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
            f'{tblLayout}'
            '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="9343"/></w:tblGrid>'
            f'{rows_xml}'
            '</w:tbl>')


def make_document(table_xml, use_d4d126_docgrid=True):
    if use_d4d126_docgrid:
        docgrid = '<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/>'
    else:
        docgrid = '<w:docGrid w:type="lines" w:linePitch="360"/>'
    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="851" w:footer="992" w:gutter="0"/>'
            f'{docgrid}'
            '</w:sectPr>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {NS}><w:body>{table_xml}<w:p/>{sect}</w:body></w:document>')


def build_docx(name, *, with_spacing, with_pstyle_ac, with_after, v_align,
               use_d4d126_docgrid, with_trHeight, tblLayout_fixed):
    """Build a 7-row table mirroring d4d126 t1.

    with_trHeight: if True, rows 2/4/6 have trHeight per d4d126.
    """
    p = OUT / f'{name}.docx'
    if p.exists(): p.unlink()
    # 7 rows; row 2/4/6 have trHeight in d4d126
    rows = []
    th_per_row = {2: 658, 4: 437, 6: 549}
    for i in range(1, 8):
        th = th_per_row.get(i) if with_trHeight else None
        rows.append(make_row(label_idx=i, trHeight=th, v_align=v_align,
                             with_spacing=with_spacing, with_pstyle_ac=with_pstyle_ac,
                             with_after=with_after))
    table_xml = make_table(''.join(rows), tblLayout_fixed=tblLayout_fixed)
    doc_xml = make_document(table_xml, use_d4d126_docgrid=use_d4d126_docgrid)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES_FULL)
        z.writestr('word/document.xml', doc_xml)
    print(f'  wrote {p}')


if __name__ == '__main__':
    print('d4d126 t1 mirror + incremental simplifications:')
    common = dict(with_spacing=True, with_pstyle_ac=True, with_after=True,
                  v_align='center', use_d4d126_docgrid=True,
                  with_trHeight=True, tblLayout_fixed=True)

    print('\n  v06: full d4d126 mirror (target — should show -2.35pt/row)')
    build_docx('v06_d4d126_mirror', **common)

    print('  v07 (control): no pStyle="ac" (Normal a only)')
    kw = dict(common); kw['with_pstyle_ac'] = False
    build_docx('v07_no_style_ac', **kw)

    print('  v08 (control): default docGrid (linePitch=360 type=lines)')
    kw = dict(common); kw['use_d4d126_docgrid'] = False
    build_docx('v08_no_d4d126_docgrid', **kw)

    print('  v09 (control): vAlign=top')
    kw = dict(common); kw['v_align'] = 'top'
    build_docx('v09_v_align_top', **kw)

    print('  v10 (control): no spacing.before/after (just line=240 exact)')
    kw = dict(common); kw['with_spacing'] = False
    build_docx('v10_no_spacing_before_after', **kw)

    print('  v11 (control): all rows auto height (no trHeight)')
    kw = dict(common); kw['with_trHeight'] = False
    build_docx('v11_no_trHeight', **kw)

    print('  v12 (control): no tblLayout=fixed (default auto)')
    kw = dict(common); kw['tblLayout_fixed'] = False
    build_docx('v12_no_tbl_layout_fixed', **kw)

    print('  v13 (control): no spacing.after (only before)')
    kw = dict(common); kw['with_after'] = False
    build_docx('v13_no_spacing_after', **kw)

    print('done.')
