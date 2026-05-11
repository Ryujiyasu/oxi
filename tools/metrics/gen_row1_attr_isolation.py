"""Day 33 part 16 — Isolate which of d4d126 row 1's 4 attributes causes
the +2.3pt Oxi over-pump.

Row 1 (long-text header) has:
  - w:vAlign w:val="center" (cell vertical center)
  - w:pStyle w:val="ac" (some style, contents undefined)
  - w:spacing w:beforeLines="30" w:before="87" w:line="240" w:lineRule="exact"
  - w:tcBorders/w:bottom w:val="dashed"

Strategy: 2x2x2x2 grid of variants but constrained to 8 to keep manageable.
Baseline (no attrs) + each attr alone + all combined.

Each repro is a 1-row 1-cell table with long text "（提供申出者が国際機関の場合は、本欄に記載する。）"
followed by an anchor paragraph below the table. row_height = anchor_y - row_y.
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'row1_attr_isolation')

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

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

# Defines docDefault sz=21 (10.5pt) MS Mincho — matches d4d126
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="ac"><w:name w:val="ittarou"/><w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:spacing w:line="210" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr><w:rPr><w:rFonts w:ascii="MS Mincho" w:hAnsi="MS Mincho"/><w:spacing w:val="-1"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>
</w:styles>"""


def make_doc(vAlign: str, pStyle: bool, spacing: bool, lineRuleExact: bool):
    """Build doc with optional attributes set on row 1 cell paragraph."""
    pPr_inner = ''
    if pStyle:
        pPr_inner += '<w:pStyle w:val="ac"/>'
    if spacing or lineRuleExact:
        spacing_xml = '<w:spacing'
        if spacing:
            spacing_xml += ' w:beforeLines="30" w:before="87"'
        if lineRuleExact:
            spacing_xml += ' w:line="240" w:lineRule="exact"'
        spacing_xml += '/>'
        pPr_inner += spacing_xml
    pPr_xml = f'<w:pPr>{pPr_inner}</w:pPr>' if pPr_inner else ''

    valign_xml = f'<w:vAlign w:val="{vAlign}"/>' if vAlign else ''
    cell_para = (
        f'<w:p>{pPr_xml}'
        f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
        f'<w:t>（提供申出者が国際機関の場合は、本欄に記載する。）</w:t></w:r>'
        f'</w:p>'
    )
    table = (
        '<w:tbl>'
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
        '<w:tr><w:tc>'
        f'<w:tcPr><w:tcW w:w="9000" w:type="dxa"/>{valign_xml}</w:tcPr>'
        f'{cell_para}'
        '</w:tc></w:tr>'
        '</w:tbl>'
    )
    anchor = '<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>下のアンカー</w:t></w:r></w:p>'
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{table}
{anchor}
<w:sectPr>
<w:pgSz w:w="11904" w:h="16836" w:code="9"/>
<w:pgMar w:top="851" w:right="1134" w:bottom="142" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720"/>
</w:sectPr>
</w:body>
</w:document>"""


def write(label: str, doc: str):
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


# 8 variants: baseline + each attr alone + all combined + pairwise
variants = [
    ('R1A_baseline',          '',       False, False, False),
    ('R1A_vAlign_center',     'center', False, False, False),
    ('R1A_pStyle_ac',          '',      True,  False, False),
    ('R1A_spacing_before',     '',      False, True,  False),
    ('R1A_lineRule_exact',     '',      False, False, True),
    ('R1A_spacing_lineRule',   '',      False, True,  True),
    ('R1A_vAlign_lineRule',    'center',False, False, True),
    ('R1A_all4',               'center',True,  True,  True),
]

for label, vAlign, pStyle, spacing, lineRule in variants:
    doc = make_doc(vAlign, pStyle, spacing, lineRule)
    write(label, doc)
print(f'Wrote {len(variants)} variants to {OUT_DIR}')
