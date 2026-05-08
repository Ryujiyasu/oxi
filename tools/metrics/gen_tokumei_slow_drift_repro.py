"""Generate minimal repros isolating the tokumei SLOW ACCUMULATION drift.

Day 31 finding: tokumei sub-family (d4d126/de6e/6514/a1d6/191cb) shows
~+0.05-0.10pt drift per cell row, accumulating linearly over 25-73 paragraphs.
The dominant cell paragraph pattern in these docs is:

  <w:pPr>
    <w:pStyle w:val="ac"/>
    <w:spacing w:beforeLines="30" w:before="87" w:afterLines="30" w:after="87"
               w:line="240" w:lineRule="exact"/>
    <w:rPr><w:spacing w:val="0"/></w:rPr>
  </w:pPr>
  + tcPr vAlign="center"

Each variant: 30 consecutive 1-row 1-cell tables with the above paragraph
pattern, varying:
  - V100: baseline (lineRule=exact line=240, before/after=87, vAlign=center)
  - V101: lineRule=exact line=240, vAlign=center, NO before/after spacing
  - V102: lineRule=exact line=240, NO vAlign, before/after=87
  - V103: lineRule=auto, vAlign=center, before/after=87 (lineRule=auto = grid behavior)
  - V104: lineRule=exact line=240, vAlign=center, before/after=87, NO adjustLineHeightInTable

Output: tools/golden-test/repros/tokumei_slow_drift/TS_<label>.docx
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'tokumei_slow_drift')

NUM_ROWS = 30


def make_para(line_rule: str, line_val: str,
              before: str | None, after: str | None) -> str:
    spacing_attrs = []
    if before:
        spacing_attrs.append(f'w:beforeLines="30" w:before="{before}"')
    if after:
        spacing_attrs.append(f'w:afterLines="30" w:after="{after}"')
    spacing_attrs.append(f'w:line="{line_val}" w:lineRule="{line_rule}"')
    spacing = '<w:spacing ' + ' '.join(spacing_attrs) + '/>'
    return ('<w:p>'
            '<w:pPr>'
            f'{spacing}'
            '<w:rPr><w:spacing w:val="0"/></w:rPr>'
            '</w:pPr>'
            '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
            f'<w:t>当該公的機関の名称{{i}}</w:t></w:r>'
            '</w:p>')


def make_row(idx: int, para_xml: str, valign: bool) -> str:
    valign_xml = '<w:vAlign w:val="center"/>' if valign else ''
    return ('<w:tbl>'
            '<w:tblPr>'
            '<w:tblW w:w="9000" w:type="dxa"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4"/>'
            '<w:left w:val="single" w:sz="4"/>'
            '<w:bottom w:val="single" w:sz="4"/>'
            '<w:right w:val="single" w:sz="4"/>'
            '</w:tblBorders>'
            '</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
            '<w:tr>'
            '<w:tc>'
            f'<w:tcPr><w:tcW w:w="9000" w:type="dxa"/>{valign_xml}</w:tcPr>'
            f'{para_xml.replace("{i}", str(idx))}'
            '</w:tc>'
            '</w:tr>'
            '</w:tbl>')


SETTINGS_WITH_FLAG = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

SETTINGS_NO_FLAG = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
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


def doc_xml(rows_xml: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{rows_xml}
<w:sectPr>
<w:pgSz w:w="11904" w:h="16836" w:code="9"/>
<w:pgMar w:top="851" w:right="1134" w:bottom="851" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720"/>
<w:docGrid w:type="linesAndChars" w:linePitch="272"/>
</w:sectPr>
</w:body>
</w:document>"""


def write_docx(label: str, doc: str, settings: str):
    out = os.path.join(OUT_DIR, f'{label}.docx')
    os.makedirs(OUT_DIR, exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', DOC_RELS)
        zf.writestr('word/settings.xml', settings)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/document.xml', doc)
    return out


variants = [
    # V100: tokumei baseline pattern
    ('TS_V100_baseline',
     make_para('exact', '240', '87', '87'), True, SETTINGS_WITH_FLAG),
    # V101: no spacing before/after — isolates lineRule=exact only
    ('TS_V101_no_before_after',
     make_para('exact', '240', None, None), True, SETTINGS_WITH_FLAG),
    # V102: no vAlign — isolates vAlign=center contribution
    ('TS_V102_no_valign',
     make_para('exact', '240', '87', '87'), False, SETTINGS_WITH_FLAG),
    # V103: lineRule=auto — different lh path
    ('TS_V103_lineRule_auto',
     make_para('auto', '240', '87', '87'), True, SETTINGS_WITH_FLAG),
    # V104: NO adjustLineHeightInTable flag
    ('TS_V104_no_flag',
     make_para('exact', '240', '87', '87'), True, SETTINGS_NO_FLAG),
    # V105: line=200 (10pt exact) — sub-default-fs-size
    ('TS_V105_line200',
     make_para('exact', '200', '87', '87'), True, SETTINGS_WITH_FLAG),
    # V106: line=300 (15pt exact) — over-default
    ('TS_V106_line300',
     make_para('exact', '300', '87', '87'), True, SETTINGS_WITH_FLAG),
]


def main():
    for label, para_xml, valign, settings in variants:
        rows = '\n'.join(make_row(i, para_xml, valign) for i in range(NUM_ROWS))
        p = write_docx(label, doc_xml(rows), settings)
        print(f'wrote {p} (rows={NUM_ROWS})')
    print('Done.')


if __name__ == '__main__':
    main()
