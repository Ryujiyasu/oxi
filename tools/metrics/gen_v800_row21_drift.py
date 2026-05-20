"""S155: V800 minimal repros to isolate a1d6 row 21's intrinsic +14pt drift.

a1d6 row 21:
- 2-cell row (narrow cell[0] tcW=1985 + wide cell[1] tcW=7654 gridSpan=8)
- cell[1] vAlign=center, 19 paragraphs (OVERFLOW)
- trHeight=4822 atLeast (= 241.1pt)
- Section header: sb=146 (beforeLines=50) line=280 exact, sz=20, hanging indent

Variants to isolate the +14pt drift cause:
  V800a: 1-cell (wide only) + 19 paras + vAlign=center + trH=4822 atLeast
  V800b: same but vAlign=top
  V800c: same but no trHeight
  V800d: same but trH=4822 exact
  V800e: 1 para (section header only, no overflow)
  V800f: 19 paras, section header no sb
  V800g: 1 para + trH=4822 atLeast (overflow not triggered)
  V800h: 2-cell row mimicking a1d6 exactly (narrow + wide)
  V800i: 1 para in 1-cell row, vAlign=center, no trH
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'a1d6_row21_isolate')

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="ac">
<w:name w:val="ac"/>
<w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:spacing w:line="210" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr>
<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="ＭＳ 明朝"/><w:spacing w:val="-1"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
</w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:useFELayout/>
<w:adjustLineHeightInTable/>
<w:doNotExpandShiftReturn/>
</w:compat>
</w:settings>'''


def section_header():
    """Mimics a1d6 i=237 ※ section header."""
    return ('<w:p><w:pPr>'
            '<w:pStyle w:val="ac"/><w:wordWrap/>'
            '<w:spacing w:beforeLines="50" w:before="146" w:line="280" w:lineRule="exact"/>'
            '<w:ind w:leftChars="50" w:left="316" w:hangingChars="100" w:hanging="207"/>'
            '<w:jc w:val="left"/>'
            '<w:rPr><w:spacing w:val="0"/><w:sz w:val="20"/></w:rPr>'
            '</w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="20"/></w:rPr>'
            '<w:t>※　匿名データを取り扱う者が以下のいずれにも該当しない場合</w:t></w:r>'
            '</w:p>')


def section_header_no_sb():
    return ('<w:p><w:pPr>'
            '<w:pStyle w:val="ac"/><w:wordWrap/>'
            '<w:spacing w:line="280" w:lineRule="exact"/>'
            '<w:rPr><w:sz w:val="20"/></w:rPr>'
            '</w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="20"/></w:rPr>'
            '<w:t>※　匿名データを取り扱う者が以下のいずれにも該当しない場合</w:t></w:r>'
            '</w:p>')


def filler_para(i):
    return ('<w:p><w:pPr>'
            '<w:pStyle w:val="ac"/>'
            '<w:spacing w:line="280" w:lineRule="exact"/>'
            '<w:rPr><w:sz w:val="20"/></w:rPr>'
            '</w:pPr>'
            f'<w:r><w:rPr><w:sz w:val="20"/></w:rPr>'
            f'<w:t>FILLER_{i}_この行は長いです覆い隠す内容</w:t></w:r>'
            '</w:p>')


def body_marker(text):
    return f'<w:p><w:pPr><w:pStyle w:val="ac"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def make_row(content, valign=None, trheight=None, trh_rule=None,
             tcw=7654, gridspan=8, narrow_cell=False):
    valign_xml = f'<w:vAlign w:val="{valign}"/>' if valign else ''
    trh_xml = ''
    if trheight is not None:
        attrs = f'w:val="{trheight}"'
        if trh_rule:
            attrs += f' w:hRule="{trh_rule}"'
        trh_xml = f'<w:trPr><w:trHeight {attrs}/></w:trPr>'
    gs_xml = f'<w:gridSpan w:val="{gridspan}"/>' if gridspan > 1 else ''
    narrow_xml = ''
    if narrow_cell:
        # Add a narrow header cell mimicking a1d6 cell[0]
        narrow_xml = ('<w:tc><w:tcPr>'
                      '<w:tcW w:w="1985" w:type="dxa"/>'
                      '</w:tcPr>'
                      '<w:p><w:pPr><w:pStyle w:val="ac"/></w:pPr></w:p>'
                      '</w:tc>')
    main_cell = ('<w:tc><w:tcPr>'
                 f'<w:tcW w:w="{tcw}" w:type="dxa"/>{gs_xml}'
                 f'{valign_xml}'
                 '</w:tcPr>'
                 f'{content}'
                 '</w:tc>')
    return f'<w:tr>{trh_xml}{narrow_xml}{main_cell}</w:tr>'


def make_table(rows_xml, total_w=7654, n_cols=8):
    col_w = total_w // n_cols
    grid_cols = ''.join(f'<w:gridCol w:w="{col_w}"/>' for _ in range(n_cols))
    return ('<w:tbl>'
            '<w:tblPr>'
            f'<w:tblW w:w="{total_w}" w:type="dxa"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
            '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>'
            '</w:tblBorders>'
            '</w:tblPr>'
            f'<w:tblGrid>{grid_cols}</w:tblGrid>'
            f'{rows_xml}'
            '</w:tbl>')


def doc_xml(body):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1247" w:right="1077" w:bottom="1440" w:left="1077" w:header="851" w:footer="992"/>
<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/>
</w:sectPr>
</w:body>
</w:document>"""


def write(label, doc):
    out = os.path.join(OUT_DIR, f'{label}.docx')
    os.makedirs(OUT_DIR, exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', DOC_RELS)
        zf.writestr('word/settings.xml', SETTINGS)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/document.xml', doc)


def main():
    # 19 paras = section + 18 filler
    n_paras_overflow = 19
    overflow_content = section_header() + ''.join(filler_para(i) for i in range(n_paras_overflow - 1))
    single_content = section_header()

    # V800a: 1-cell wide + 19 paras + vAlign=center + trH=4822 atLeast
    write('V800a_overflow_valign_trH_atLeast', doc_xml(make_table(
        make_row(overflow_content, valign='center', trheight=4822))))

    # V800b: vAlign=top instead of center
    write('V800b_overflow_valign_top', doc_xml(make_table(
        make_row(overflow_content, valign='top', trheight=4822))))

    # V800c: no trHeight (auto)
    write('V800c_overflow_no_trH', doc_xml(make_table(
        make_row(overflow_content, valign='center'))))

    # V800d: trH=4822 EXACT (not atLeast)
    write('V800d_overflow_trH_exact', doc_xml(make_table(
        make_row(overflow_content, valign='center', trheight=4822, trh_rule='exact'))))

    # V800e: only section header (1 para, no overflow)
    write('V800e_single_para_trH', doc_xml(make_table(
        make_row(single_content, valign='center', trheight=4822))))

    # V800f: 19 paras + section header WITHOUT sb
    content_no_sb = section_header_no_sb() + ''.join(filler_para(i) for i in range(n_paras_overflow - 1))
    write('V800f_overflow_no_sb', doc_xml(make_table(
        make_row(content_no_sb, valign='center', trheight=4822))))

    # V800g: 1 para + trH=4822 atLeast (overflow not triggered)
    write('V800g_1para_trH_atLeast', doc_xml(make_table(
        make_row(single_content, valign='center', trheight=4822))))

    # V800h: 2-cell row mimicking a1d6 exactly (narrow header + wide content)
    write('V800h_2cell_a1d6_exact', doc_xml(make_table(
        make_row(overflow_content, valign='center', trheight=4822, narrow_cell=True))))

    # V800i: 1 para in 1-cell, vAlign=center, no trH (control)
    write('V800i_1para_no_trH', doc_xml(make_table(
        make_row(single_content, valign='center'))))

    print('Done. Repros in', OUT_DIR)


if __name__ == '__main__':
    main()
