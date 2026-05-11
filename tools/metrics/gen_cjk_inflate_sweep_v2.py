"""Day 33 part 13 — extend cjk_inflate_sweep to multiple CJK 83/64 fonts.

Same structure as gen_cjk_inflate_sweep.py but parameterised by font family.
Generates 1-row 1-cell tables with 2 paragraphs (row gap = line height).
Goal: verify whether Word's cell line height formula
  round(font_size * (winA+winD)/UPM * 83/64, 0.5pt)
holds for fonts with UPM=2048 (Yu Mincho, Yu Gothic) as well as UPM=256
(MS Mincho, MS Gothic). MS Mincho fs={8,9,10,10.5,11,12,14} already verified
2026-05-08 (pipeline_data/cjk_inflate_oxi.json).

Output: tools/golden-test/repros/cjk_inflate_v2/CJK2_<font>_fs<N>_snap0.docx
"""
from __future__ import annotations
import os, zipfile

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DIR = os.path.join(REPO, 'tools', 'golden-test', 'repros', 'cjk_inflate_v2')

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

def styles(font: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
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


def make_para(font: str, fs_halfpt: int, text: str) -> str:
    return (
        f'<w:p>'
        f'<w:pPr><w:snapToGrid w:val="0"/>'
        f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}"/>'
        f'<w:sz w:val="{fs_halfpt}"/><w:szCs w:val="{fs_halfpt}"/></w:rPr>'
        f'</w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hint="eastAsia"/>'
        f'<w:sz w:val="{fs_halfpt}"/></w:rPr>'
        f'<w:t>{text}</w:t></w:r>'
        f'</w:p>'
    )


def make_doc(font: str, fs_halfpt: int) -> str:
    p1 = make_para(font, fs_halfpt, 'テスト１')
    p2 = make_para(font, fs_halfpt, 'テスト２')
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
        '<w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>'
        f'{p1}{p2}'
        '</w:tc></w:tr>'
        '</w:tbl>'
    )
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


def write(label: str, font: str, doc: str):
    out = os.path.join(OUT_DIR, f'{label}.docx')
    os.makedirs(OUT_DIR, exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', DOC_RELS)
        zf.writestr('word/settings.xml', SETTINGS)
        zf.writestr('word/styles.xml', styles(font))
        zf.writestr('word/document.xml', doc)
    return out


# Fonts: 4 distinct CJK 83/64 families (Meiryo skipped since no font_metrics entry)
fonts = [
    ('MS Mincho', 'MSM'),
    ('MS Gothic', 'MSG'),
    ('Yu Mincho', 'YuM'),
    ('Yu Gothic', 'YuG'),
]
font_sizes = [(8, 16), (9, 18), (10, 20), (10.5, 21), (11, 22), (12, 24), (14, 28)]

for full_name, short in fonts:
    for fs_pt, fs_hp in font_sizes:
        label = f'CJK2_{short}_fs{str(fs_pt).replace(".","p")}_snap0'
        write(label, full_name, make_doc(full_name, fs_hp))

print(f'Wrote {len(fonts) * len(font_sizes)} docx to {OUT_DIR}')
