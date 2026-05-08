"""Generate minimal repros isolating the post-table cursor_y over-advance bug
suspected in 1636 pagination FAIL.

Each variant: A4 page, top=42.55pt, bottom=7.10pt, NO footer, with 1 table near
body bottom + 1 'bibou-like' paragraph after the table. Vary the trailing-cell
structure to localize what makes Oxi push the post-paragraph to page 2.

Output: tools/golden-test/repros/post_tbl_cursor/PT_<label>.docx
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'post_tbl_cursor')

# Spacers — 18 empty paras 30pt each ≈ 540pt → table starts near y=580
SPACER = '<w:p><w:pPr><w:spacing w:line="600" w:lineRule="exact"/></w:pPr></w:p>' * 18

# Tighter spacer for "table near body bottom" variants — push table down
# to leave only ~30pt below table for bibou
SPACER_TIGHT = '<w:p><w:pPr><w:spacing w:line="600" w:lineRule="exact"/></w:pPr></w:p>' * 23

# Bibou-like paragraph after table (matches 1636 i=79 structure)
BIBOU = ('<w:p>'
         '<w:pPr>'
         '<w:snapToGrid w:val="0"/>'
         '<w:spacing w:beforeLines="100" w:before="272" w:line="240" w:lineRule="auto"/>'
         '<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
         '</w:pPr>'
         '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
         '<w:t>備考</w:t></w:r>'
         '</w:p>')

# Settings: include adjustLineHeightInTable like 1636
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


def make_table(cell_inner_xml: str) -> str:
    """1-row 1-col table with given cell content."""
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
            '<w:tr><w:tc>'
            '<w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>'
            f'{cell_inner_xml}'
            '</w:tc></w:tr>'
            '</w:tbl>')


def doc_xml(table_xml: str, spacer: str = SPACER) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{spacer}
{table_xml}
{BIBOU}
<w:sectPr>
<w:pgSz w:w="11904" w:h="16836" w:code="9"/>
<w:pgMar w:top="851" w:right="1134" w:bottom="142" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>
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


# Variants
# PT_V1: simplest — 1 cell with 1 line of text only (like '所定の金額の')
v1_cell = ('<w:p><w:pPr><w:spacing w:line="200" w:lineRule="exact"/></w:pPr>'
           '<w:r><w:rPr><w:sz w:val="16"/></w:rPr>'
           '<w:t>所定の金額の収入印紙を貼り、消印しないこと</w:t></w:r></w:p>')

# PT_V2: V1 + widow-control empty para (like 1636 cell middle)
v2_cell = v1_cell + (
    '<w:p><w:pPr><w:widowControl/>'
    '<w:ind w:left="210" w:rightChars="95" w:right="199" w:hangingChars="100" w:hanging="210"/>'
    '<w:rPr><w:rFonts w:ascii="Century" w:hAnsi="Century"/><w:szCs w:val="22"/></w:rPr>'
    '</w:pPr></w:p>')

# PT_V3: V2 + 2 trailing snapToGrid=0 line=259/280 exact empty paras (like 1636 cell tail)
v3_cell = v2_cell + (
    '<w:p><w:pPr><w:snapToGrid w:val="0"/>'
    '<w:spacing w:line="259" w:lineRule="exact"/>'
    '<w:rPr><w:color w:val="000000"/></w:rPr>'
    '</w:pPr></w:p>'
    '<w:p><w:pPr><w:snapToGrid w:val="0"/>'
    '<w:spacing w:line="280" w:lineRule="exact"/>'
    '<w:rPr><w:color w:val="000000"/></w:rPr>'
    '</w:pPr></w:p>')

# PT_V4: V3 with adjustLineHeightInTable OFF
# PT_V5: V1 with adjustLineHeightInTable OFF  (control)
# PT_V6: V2 with adjustLineHeightInTable OFF

variants = [
    ('PT_V1_simple_with_flag', v1_cell, SETTINGS_WITH_FLAG, SPACER),
    ('PT_V2_with_widowctl_empty', v2_cell, SETTINGS_WITH_FLAG, SPACER),
    ('PT_V3_with_3trailing_empties', v3_cell, SETTINGS_WITH_FLAG, SPACER),
    ('PT_V4_3trailing_NO_flag', v3_cell, SETTINGS_NO_FLAG, SPACER),
    ('PT_V5_simple_NO_flag', v1_cell, SETTINGS_NO_FLAG, SPACER),
    # Tight: push table close to body bottom (~30pt below for bibou)
    ('PT_V6_tight_simple_with_flag', v1_cell, SETTINGS_WITH_FLAG, SPACER_TIGHT),
    ('PT_V7_tight_3trailing_with_flag', v3_cell, SETTINGS_WITH_FLAG, SPACER_TIGHT),
    ('PT_V8_tight_3trailing_NO_flag', v3_cell, SETTINGS_NO_FLAG, SPACER_TIGHT),
]

for label, cell, settings, spacer in variants:
    p = write_docx(label, doc_xml(make_table(cell), spacer), settings)
    print(f'wrote {p}')
print('Done.')
