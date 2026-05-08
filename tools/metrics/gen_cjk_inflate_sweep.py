"""Generate minimal repros for CJK 83/64 inflate threshold investigation.

For each (font_size, snap_to_grid) combination:
- 1-row 1-cell table with 2 paragraphs (so we can measure row gap = line height).
- Cell paragraphs: MS Mincho, sz=N, snap=S, in cell with adjustLineHeightInTable=true.
- Measure Word's row gap = actual rendered line height.

Output: tools/golden-test/repros/cjk_inflate/CI_<fs>pt_snap<S>.docx (14 variants).

Goal: find which (fs, snap) combos Word renders at the small line height
(~natural lh) vs the inflated 83/64 line height.
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'cjk_inflate')

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


def make_para(fs_halfpt: int, snap: int, text: str) -> str:
    """One paragraph in cell: sz=fs_halfpt(*0.5pt), snap=snap, text."""
    snap_xml = f'<w:snapToGrid w:val="{snap}"/>'
    return (
        f'<w:p>'
        f'<w:pPr>{snap_xml}'
        f'<w:rPr><w:sz w:val="{fs_halfpt}"/><w:szCs w:val="{fs_halfpt}"/></w:rPr>'
        f'</w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="{fs_halfpt}"/></w:rPr>'
        f'<w:t>{text}</w:t></w:r>'
        f'</w:p>'
    )


def make_table(fs_halfpt: int, snap: int, valign: str = '') -> str:
    """1-row 1-cell table with 2 same-format paragraphs."""
    p1 = make_para(fs_halfpt, snap, 'テスト１')
    p2 = make_para(fs_halfpt, snap, 'テスト２')
    valign_xml = f'<w:vAlign w:val="{valign}"/>' if valign else ''
    return (
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
        f'{p1}{p2}'
        '</w:tc></w:tr>'
        '</w:tbl>'
    )


def make_doc(fs_halfpt: int, snap: int, valign: str = '') -> str:
    table = make_table(fs_halfpt, snap, valign)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{table}
<w:p/>
<w:sectPr>
<w:pgSz w:w="11904" w:h="16836" w:code="9"/>
<w:pgMar w:top="851" w:right="1134" w:bottom="142" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720"/>
<w:docGrid w:type="linesAndChars" w:linePitch="272"/>
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


# Font sizes (full pt × 2 = halfpt for w:sz)
font_sizes = [(8, 16), (9, 18), (10, 20), (10.5, 21), (11, 22), (12, 24), (14, 28)]
snap_modes = [0, 1]

for fs_pt, fs_hp in font_sizes:
    for snap in snap_modes:
        label = f'CI_fs{str(fs_pt).replace(".","p")}_snap{snap}'
        p = write(label, make_doc(fs_hp, snap))
        print(f'wrote {p}')

# Add vAlign=center variants for 1636 minimal repro
for fs_pt, fs_hp in [(8, 16), (10.5, 21)]:
    for snap in [0, 1]:
        label = f'CI_fs{str(fs_pt).replace(".","p")}_snap{snap}_valignCenter'
        p = write(label, make_doc(fs_hp, snap, 'center'))
        print(f'wrote {p}')

print('Done.')
