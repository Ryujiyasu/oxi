"""S152: V500 cell-context empty-para + section header repros.

a1d6 pattern at i=237 (取り扱う):
- Inside a cell
- 5 empty paragraphs with pStyle="ac" + line=280 exact precede the section
- Section header has sb=146 beforeLines=50, line=280 exact, sz=20, hanging indent

Drift in a1d6: +11.55pt (Oxi places section header later than Word).

V400 (body context) showed pStyle="ac" empty paras add only +0.05pt drift.
Cell context must differ.

Variants:
  V500a: cell with 0 empty paras + section header (control)
  V500b: cell with 1 empty `pStyle="ac" line=280 exact` + section header
  V500c: cell with 5 empty same → mimics a1d6 i=237 setup
  V500d: same as V500c but with vAlign=center on cell
  V500e: same as V500c but trHeight unset (auto)
  V500f: same as V500c but with explicit trHeight large enough
  V500g: section header WITHOUT sb (no beforeLines/before)
  V500h: 5 BARE `<w:p/>` empty paras (default) + section
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'cell_empty_section')

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


def empty_para_ac_exact280():
    """Empty para with pStyle="ac" + line=280 exact (a1d6 pattern)."""
    return ('<w:p><w:pPr><w:pStyle w:val="ac"/>'
            '<w:spacing w:line="280" w:lineRule="exact"/>'
            '<w:jc w:val="center"/>'
            '</w:pPr></w:p>')


def empty_para_default():
    return '<w:p/>'


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
    """Section header without spacing-before."""
    return ('<w:p><w:pPr>'
            '<w:pStyle w:val="ac"/><w:wordWrap/>'
            '<w:spacing w:line="280" w:lineRule="exact"/>'
            '<w:rPr><w:sz w:val="20"/></w:rPr>'
            '</w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="20"/></w:rPr>'
            '<w:t>※　SECTION_NO_SB</w:t></w:r>'
            '</w:p>')


def body_marker(text):
    return f'<w:p><w:pPr><w:pStyle w:val="ac"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def make_cell(content: str, valign: str | None = None, trheight: int | None = None,
              trh_rule: str | None = None) -> str:
    valign_xml = f'<w:vAlign w:val="{valign}"/>' if valign else ''
    trh_xml = ''
    if trheight is not None:
        attrs = f'w:val="{trheight}"'
        if trh_rule:
            attrs += f' w:hRule="{trh_rule}"'
        trh_xml = f'<w:trPr><w:trHeight {attrs}/></w:trPr>'
    return ('<w:tbl>'
            '<w:tblPr>'
            '<w:tblW w:w="9000" w:type="dxa"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
            '<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>'
            '</w:tblBorders>'
            '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
            f'<w:tr>{trh_xml}'
            '<w:tc><w:tcPr>'
            '<w:tcW w:w="9000" w:type="dxa"/>'
            f'{valign_xml}'
            '</w:tcPr>'
            f'{content}'
            '</w:tc></w:tr></w:tbl>')


def doc_xml(body: str) -> str:
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


def write_docx(label: str, doc: str):
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
    before = body_marker('BEFORE')
    after = body_marker('AFTER')

    # V500a: cell with section header only (no empty)
    cell_a = make_cell(section_header())
    write_docx('V500a_cell_section_only', doc_xml(before + cell_a + after))

    # V500b: cell with 1 empty + section
    cell_b = make_cell(empty_para_ac_exact280() + section_header())
    write_docx('V500b_cell_1empty_section', doc_xml(before + cell_b + after))

    # V500c: cell with 5 empty + section (a1d6 pattern)
    cell_c = make_cell(empty_para_ac_exact280() * 5 + section_header())
    write_docx('V500c_cell_5empty_section', doc_xml(before + cell_c + after))

    # V500d: same as V500c but vAlign=center
    cell_d = make_cell(empty_para_ac_exact280() * 5 + section_header(), valign='center')
    write_docx('V500d_cell_5empty_section_valign_center', doc_xml(before + cell_d + after))

    # V500e: cell with 5 empty + section, explicit trHeight=200 (small, content forces grow)
    cell_e = make_cell(empty_para_ac_exact280() * 5 + section_header(), trheight=200)
    write_docx('V500e_cell_5empty_section_trheight200', doc_xml(before + cell_e + after))

    # V500f: cell with 5 empty + section, explicit trHeight=3000 (large, atLeast)
    cell_f = make_cell(empty_para_ac_exact280() * 5 + section_header(), trheight=3000)
    write_docx('V500f_cell_5empty_section_trheight3000', doc_xml(before + cell_f + after))

    # V500g: cell with 5 empty + section_header WITHOUT sb
    cell_g = make_cell(empty_para_ac_exact280() * 5 + section_header_no_sb())
    write_docx('V500g_cell_5empty_section_no_sb', doc_xml(before + cell_g + after))

    # V500h: 5 BARE empty default + section
    cell_h = make_cell(empty_para_default() * 5 + section_header())
    write_docx('V500h_cell_5bare_empty_section', doc_xml(before + cell_h + after))

    # V500i: 1 empty + section (sanity for 1-para case in cell)
    cell_i = make_cell(empty_para_ac_exact280() + section_header())
    write_docx('V500i_cell_1empty_ac_section', doc_xml(before + cell_i + after))

    print('Done. Repros in', OUT_DIR)


if __name__ == '__main__':
    main()
