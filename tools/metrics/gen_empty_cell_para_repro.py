"""Day 33 part 15 — Minimal repro: 1-row 2-cell table where cell 1 has
N empty paragraphs (no pPr/rPr/sz, no run/sz). Vary docDefault rPr/sz
to find what controls Word's empty-cell-paragraph line height.

Variants:
- EC_def21_n2_msm: docDefault sz=21 (10.5pt), MS Mincho, 2 empty paras
- EC_def21_n3_msm: docDefault sz=21, 3 empty paras
- EC_def20_n2_msm: docDefault sz=20 (10pt), 2 empty paras
- EC_def16_n2_msm: docDefault sz=16 (8pt), 2 empty paras
- EC_nodef_n2_msm: NO docDefault sz, 2 empty paras (664c38 regime)
- EC_def21_n2_century: docDefault Century (Latin) sz=21
- EC_def21_n2_pstyle_normal_sz16: pStyle Normal sets sz=16 (8pt) for paragraph
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'empty_cell_para')

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


def styles_xml(default_sz=None, default_font='MS Mincho', extra_normal_rpr=''):
    sz_xml = f'<w:sz w:val="{default_sz}"/>' if default_sz is not None else ''
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="{default_font}" w:eastAsia="{default_font}" w:hAnsi="{default_font}"/>{sz_xml}</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>{extra_normal_rpr}</w:style>
</w:styles>"""


def doc_xml(n_empty: int):
    """Table with row of:
      cell 0 (3000tw): 1 text para "当該公的機関の名称" (default fs)
      cell 1 (6000tw): N empty paragraphs (no pPr/rPr, no run)
    Then below: anchor text para to measure y of next-row distance.
    """
    cell0_p = '<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>テキスト</w:t></w:r></w:p>'
    cell1_paras = '<w:p/>' * n_empty
    table = (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="9000" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4"/>'
        '<w:left w:val="single" w:sz="4"/>'
        '<w:bottom w:val="single" w:sz="4"/>'
        '<w:right w:val="single" w:sz="4"/>'
        '<w:insideH w:val="single" w:sz="4"/>'
        '<w:insideV w:val="single" w:sz="4"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6000"/></w:tblGrid>'
        '<w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
        + cell0_p +
        '</w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>'
        + cell1_paras +
        '</w:tc>'
        '</w:tr>'
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


def write(label, styles_x, doc_x):
    out = os.path.join(OUT_DIR, f'{label}.docx')
    os.makedirs(OUT_DIR, exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', DOC_RELS)
        zf.writestr('word/settings.xml', SETTINGS)
        zf.writestr('word/styles.xml', styles_x)
        zf.writestr('word/document.xml', doc_x)
    return out


# Variants
variants = [
    # (label, default_sz, default_font, n_empty, extra_normal_rpr)
    ('EC_def21_n2_msm', 21, 'MS Mincho', 2, ''),
    ('EC_def21_n3_msm', 21, 'MS Mincho', 3, ''),
    ('EC_def21_n1_msm', 21, 'MS Mincho', 1, ''),
    ('EC_def20_n2_msm', 20, 'MS Mincho', 2, ''),
    ('EC_def16_n2_msm', 16, 'MS Mincho', 2, ''),  # 8pt
    ('EC_def28_n2_msm', 28, 'MS Mincho', 2, ''),  # 14pt
    ('EC_nodef_n2_msm', None, 'MS Mincho', 2, ''),  # 664c38 regime
    ('EC_def21_n2_century', 21, 'Century', 2, ''),
    ('EC_def21_n2_msg', 21, 'MS Gothic', 2, ''),
    # Normal style override sz
    ('EC_def21_n2_msm_normal_sz16', 21, 'MS Mincho', 2,
     '<w:rPr><w:sz w:val="16"/></w:rPr>'),
]

for label, sz, font, n, extra in variants:
    s = styles_xml(default_sz=sz, default_font=font, extra_normal_rpr=extra)
    d = doc_xml(n)
    write(label, s, d)
print(f'Wrote {len(variants)} variants to {OUT_DIR}')
