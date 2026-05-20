"""S136: Generate TR_V200/V201/V202 to isolate the role of space_before vs
space_after vs trHeight on the tokumei row-drift bug.

Goal: verify Hypothesis H1 from `docs/design/tokumei_row_drift_fix.md`:
`is_first_block_est` subtraction at mod.rs:6302-6317 subtracts
space_before from row_height; if H1 is correct, before-only triggers
the drift, after-only does NOT, and explicit trHeight prevents the drift.

Variants (30 separate 1-row tables each, line=240 exact, vAlign=center,
border on, no explicit pad_t — same scaffold as V101):

  V200: before=87 ONLY (no after).  Predicted drift: -4.35pt/row (= -130pt cum)
        if H1 correct.
  V201: after=87 ONLY (no before).  Predicted drift: ~0 if H1 correct
        (after applies between rows, no special "first cell" handling).
  V202: before=87 + explicit trHeight=300tw (= 15pt > content+before).
        Predicted drift: ~0 because trHeight forces row to 15pt regardless
        of estimated content_h.

Sanity: V101 (no before, no after) baseline +1.0pt/row stays unchanged.

Note: this is RESEARCH only, no code changes. We're falsifying H1.
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'tokumei_slow_drift')

NUM_ROWS = 30


def make_para(line_rule: str | None, line_val: str | None,
              before: str | None = None, after: str | None = None) -> str:
    spacing_attrs = []
    if before:
        spacing_attrs.append(f'w:beforeLines="30" w:before="{before}"')
    if after:
        spacing_attrs.append(f'w:afterLines="30" w:after="{after}"')
    if line_rule and line_val:
        spacing_attrs.append(f'w:line="{line_val}" w:lineRule="{line_rule}"')
    if spacing_attrs:
        spacing = '<w:spacing ' + ' '.join(spacing_attrs) + '/>'
    else:
        spacing = ''
    return ('<w:p>'
            '<w:pPr>'
            f'{spacing}'
            '<w:rPr><w:spacing w:val="0"/></w:rPr>'
            '</w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            f'<w:t>当該公的機関の名称{{i}}</w:t></w:r>'
            '</w:p>')


def make_cell(idx: int, para_xml: str) -> str:
    return ('<w:tc>'
            '<w:tcPr><w:tcW w:w="9000" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
            f'{para_xml.replace("{i}", str(idx))}'
            '</w:tc>')


def make_single_row_table(idx: int, para_xml: str, trheight: int | None = None) -> str:
    border_xml = ('<w:tblBorders>'
                  '<w:top w:val="single" w:sz="4"/>'
                  '<w:left w:val="single" w:sz="4"/>'
                  '<w:bottom w:val="single" w:sz="4"/>'
                  '<w:right w:val="single" w:sz="4"/>'
                  '<w:insideH w:val="single" w:sz="4"/>'
                  '<w:insideV w:val="single" w:sz="4"/>'
                  '</w:tblBorders>')
    cell = make_cell(idx, para_xml)
    trh_xml = ''
    if trheight is not None:
        trh_xml = f'<w:trPr><w:trHeight w:val="{trheight}"/></w:trPr>'
    return ('<w:tbl>'
            '<w:tblPr>'
            '<w:tblW w:w="9000" w:type="dxa"/>'
            f'{border_xml}'
            '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
            f'<w:tr>{trh_xml}{cell}</w:tr>'
            '</w:tbl>')


SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""


def doc_xml(body: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{body}
<w:sectPr>
<w:pgSz w:w="11904" w:h="16836" w:code="9"/>
<w:pgMar w:top="851" w:right="1134" w:bottom="851" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720"/>
<w:docGrid w:type="linesAndChars" w:linePitch="272"/>
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
    return out


def main():
    para_before_only = make_para('exact', '240', '87', None)
    para_after_only = make_para('exact', '240', None, '87')
    para_both = make_para('exact', '240', '87', '87')

    # TR_V200: 30 separate 1-row tables, before=87 only
    tables_v200 = '\n'.join(
        make_single_row_table(i, para_before_only) for i in range(NUM_ROWS)
    )
    write_docx('TR_V200_before_only', doc_xml(tables_v200))

    # TR_V201: 30 separate 1-row tables, after=87 only
    tables_v201 = '\n'.join(
        make_single_row_table(i, para_after_only) for i in range(NUM_ROWS)
    )
    write_docx('TR_V201_after_only', doc_xml(tables_v201))

    # TR_V202: 30 separate 1-row tables, before=87 + explicit trHeight=300tw
    tables_v202 = '\n'.join(
        make_single_row_table(i, para_before_only, trheight=300) for i in range(NUM_ROWS)
    )
    write_docx('TR_V202_before_with_trheight', doc_xml(tables_v202))

    # TR_V203: 30 separate 1-row tables, before+after (same as V100 but in this gen)
    # for sanity that this generator's V100-like setup matches measured V100.
    tables_v203 = '\n'.join(
        make_single_row_table(i, para_both) for i in range(NUM_ROWS)
    )
    write_docx('TR_V203_both_sanity', doc_xml(tables_v203))

    print('Done.')


if __name__ == '__main__':
    main()
