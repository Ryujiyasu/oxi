"""Generate V107-V112 to isolate which of bugs A/B/C dominates the
tokumei slow accumulation drift.

Setup baseline (V101 from gen_tokumei_slow_drift_repro.py):
  30 separate 1-row 1-cell tables, lineRule=exact line=240, no before/after,
  vAlign=center, table border on. Result: +1.00pt/row drift.

Isolation variants:
  V107: ONE table with 30 rows (single table, multi-row) → tests if bug
        accumulates per-row WITHIN one table or only per-table.
  V108: 6 tables with 5 rows each (mixed) → tests scaling.
  V109: 30 separate tables, NO border at all (table.style.border=false) →
        tests if top_bw add is the cause (top_bw=0 if no border).
  V110: 30 separate tables, EXPLICIT pad_t=100tw on each cell → tests if
        implicit pad_t = bw is the cause (explicit overrides implicit).
  V111: 30 separate tables, lineRule UNSET (default = single/auto) →
        tests bottom-align text_y_off (only fires for exact rule).
  V112: 30 paragraphs (NO TABLES, body paragraphs) with lineRule=exact →
        tests if exact lineRule alone over-pumps (text_y_off bottom-align bug C).
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


def make_cell(idx: int, para_xml: str, valign: bool, explicit_pad_t: bool) -> str:
    valign_xml = '<w:vAlign w:val="center"/>' if valign else ''
    margin_xml = ''
    if explicit_pad_t:
        margin_xml = ('<w:tcMar>'
                      '<w:top w:w="100" w:type="dxa"/>'
                      '<w:bottom w:w="0" w:type="dxa"/>'
                      '</w:tcMar>')
    return ('<w:tc>'
            f'<w:tcPr><w:tcW w:w="9000" w:type="dxa"/>{valign_xml}{margin_xml}</w:tcPr>'
            f'{para_xml.replace("{i}", str(idx))}'
            '</w:tc>')


def make_single_row_table(idx: int, para_xml: str, valign: bool,
                          border: bool, explicit_pad_t: bool) -> str:
    border_xml = ''
    if border:
        border_xml = ('<w:tblBorders>'
                      '<w:top w:val="single" w:sz="4"/>'
                      '<w:left w:val="single" w:sz="4"/>'
                      '<w:bottom w:val="single" w:sz="4"/>'
                      '<w:right w:val="single" w:sz="4"/>'
                      '<w:insideH w:val="single" w:sz="4"/>'
                      '<w:insideV w:val="single" w:sz="4"/>'
                      '</w:tblBorders>')
    cell = make_cell(idx, para_xml, valign, explicit_pad_t)
    return ('<w:tbl>'
            '<w:tblPr>'
            '<w:tblW w:w="9000" w:type="dxa"/>'
            f'{border_xml}'
            '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
            f'<w:tr>{cell}</w:tr>'
            '</w:tbl>')


def make_multi_row_table(num_rows: int, para_xml: str, valign: bool,
                         border: bool, explicit_pad_t: bool) -> str:
    border_xml = ''
    if border:
        border_xml = ('<w:tblBorders>'
                      '<w:top w:val="single" w:sz="4"/>'
                      '<w:left w:val="single" w:sz="4"/>'
                      '<w:bottom w:val="single" w:sz="4"/>'
                      '<w:right w:val="single" w:sz="4"/>'
                      '<w:insideH w:val="single" w:sz="4"/>'
                      '<w:insideV w:val="single" w:sz="4"/>'
                      '</w:tblBorders>')
    rows = []
    for i in range(num_rows):
        cell = make_cell(i, para_xml, valign, explicit_pad_t)
        rows.append(f'<w:tr>{cell}</w:tr>')
    rows_xml = ''.join(rows)
    return ('<w:tbl>'
            '<w:tblPr>'
            '<w:tblW w:w="9000" w:type="dxa"/>'
            f'{border_xml}'
            '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
            f'{rows_xml}'
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
    para_exact = make_para('exact', '240', None, None)
    para_default = make_para(None, None, None, None)

    # V107: 1 table with 30 rows (multi-row), lineRule=exact, vAlign=center, border on, no explicit pad_t
    body_v107 = make_multi_row_table(NUM_ROWS, para_exact, True, True, False)
    write_docx('TS_V107_one_table_30rows', doc_xml(body_v107))

    # V108: 6 tables with 5 rows each
    tables_v108 = '\n'.join(
        make_multi_row_table(5, para_exact, True, True, False) for _ in range(6)
    )
    write_docx('TS_V108_6tables_5rows', doc_xml(tables_v108))

    # V109: 30 separate 1-row tables, NO BORDER (table.style.border=false → top_bw=0)
    tables_v109 = '\n'.join(
        make_single_row_table(i, para_exact, True, False, False) for i in range(NUM_ROWS)
    )
    write_docx('TS_V109_no_border', doc_xml(tables_v109))

    # V110: 30 separate 1-row tables, BORDER on, EXPLICIT pad_t=100tw
    tables_v110 = '\n'.join(
        make_single_row_table(i, para_exact, True, True, True) for i in range(NUM_ROWS)
    )
    write_docx('TS_V110_explicit_pad_t', doc_xml(tables_v110))

    # V111: 30 separate 1-row tables, lineRule UNSET (single/auto)
    tables_v111 = '\n'.join(
        make_single_row_table(i, para_default, True, True, False) for i in range(NUM_ROWS)
    )
    write_docx('TS_V111_lineRule_default', doc_xml(tables_v111))

    # V112: 30 body paragraphs (no tables) with lineRule=exact
    paras_v112 = '\n'.join(
        para_exact.replace('{i}', str(i)) for i in range(NUM_ROWS)
    )
    write_docx('TS_V112_body_paragraphs', doc_xml(paras_v112))

    print('Done.')


if __name__ == '__main__':
    main()
