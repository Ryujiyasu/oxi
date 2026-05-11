"""Day 33 part 26 (2026-05-11) — Word line-height measurement campaign generator.

Generates a matrix of minimal repros varying:
  - font (MS Mincho, MS Gothic, Yu Mincho, Yu Gothic)
  - fs (8, 9, 10.5, 11, 12, 14)
  - lhRule (auto, exact, atLeast, multiple) — represented as OOXML lineRule
  - line value (240=Single, 360, 480, custom)
  - docGrid type (none, lines, linesAndChars) × linePitch (240, 280, 312, 360)

Pairs with measure_lh_campaign.py to produce
pipeline_data/lh_campaign.json — a lookup of Word's actual per-line
advance for each combo. Future fix can replace Oxi's line_height
calculation with this table for non-trivial combos.

Initial focus (Day 33 part 24 lead): fs=10.5 + docGrid linesAndChars +
linePitch=360tw (db9ca uses this; Oxi computes 18pt, Word renders 19pt
per-line).
"""
from __future__ import annotations
import os, zipfile
from pathlib import Path

OUT = Path('tools/golden-test/repros/lh_campaign')
OUT.mkdir(parents=True, exist_ok=True)

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>'''


def make_para(font, fs_half, line_val, line_rule, snap_to_grid=True):
    """fs_half is sz half-points (Word's sz convention). e.g. fs=10.5 => sz=21."""
    spacing = ''
    if line_val:
        spacing = f'<w:spacing w:line="{line_val}" w:lineRule="{line_rule}"/>'
    snap = '' if snap_to_grid else '<w:snapToGrid w:val="0"/>'
    rfonts = f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
    # Use Latin text for non-CJK fonts (Times New Roman, Century, etc.)
    is_latin = font in ('Times New Roman', 'Century', 'Arial', 'Calibri')
    test_text = 'Test line content' if is_latin else 'テスト行'
    return ('<w:p><w:pPr>'
            + spacing + snap +
            f'<w:rPr>{rfonts}<w:sz w:val="{fs_half}"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rfonts}<w:sz w:val="{fs_half}"/></w:rPr>'
            f'<w:t>{test_text}</w:t></w:r></w:p>')


def make_document(font, fs_half, line_val, line_rule, grid_type, line_pitch, snap):
    # 5 short paragraphs to measure per-line advance
    paras = ''.join(make_para(font, fs_half, line_val, line_rule, snap) for _ in range(5))
    grid_xml = ''
    if grid_type != 'none':
        grid_xml = f'<w:docGrid w:type="{grid_type}" w:linePitch="{line_pitch}"/>'
    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397" w:gutter="0"/>'
            + grid_xml +
            '</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{paras}{sect}</w:body></w:document>'


def build_docx(name, font, fs_half, line_val, line_rule, grid_type, line_pitch, snap):
    p = OUT / f'{name}.docx'
    if p.exists(): p.unlink()
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/document.xml',
                   make_document(font, fs_half, line_val, line_rule, grid_type, line_pitch, snap))
    print(f'  wrote {p}')


# Day 33 part 24 specific case: fs=10.5 MS Mincho, docGrid linesAndChars
# linePitch=360tw — db9ca actual setup. Word: 19pt/line, Oxi: 18pt/line.
TEST_MATRIX = [
    # (name, font, fs_half, line_val, line_rule, grid_type, line_pitch, snap)
    # === Day 33 part 24 specific case ===
    ('LH_db9ca_replica', 'ＭＳ 明朝', 21, '', 'auto', 'linesAndChars', '360', True),

    # === Vary font ===
    ('LH_msMincho_10p5_g360',  'ＭＳ 明朝',     21, '', 'auto', 'linesAndChars', '360', True),
    ('LH_msGothic_10p5_g360',  'ＭＳ ゴシック',  21, '', 'auto', 'linesAndChars', '360', True),
    ('LH_yuMincho_10p5_g360',  '游明朝',         21, '', 'auto', 'linesAndChars', '360', True),
    ('LH_yuGothic_10p5_g360',  '游ゴシック',      21, '', 'auto', 'linesAndChars', '360', True),

    # === Vary fs ===
    ('LH_msMincho_8_g360',    'ＭＳ 明朝', 16, '', 'auto', 'linesAndChars', '360', True),
    ('LH_msMincho_9_g360',    'ＭＳ 明朝', 18, '', 'auto', 'linesAndChars', '360', True),
    ('LH_msMincho_10p5_g360', 'ＭＳ 明朝', 21, '', 'auto', 'linesAndChars', '360', True),
    ('LH_msMincho_11_g360',   'ＭＳ 明朝', 22, '', 'auto', 'linesAndChars', '360', True),
    ('LH_msMincho_12_g360',   'ＭＳ 明朝', 24, '', 'auto', 'linesAndChars', '360', True),

    # === Vary linePitch ===
    ('LH_msMincho_10p5_g240', 'ＭＳ 明朝', 21, '', 'auto', 'linesAndChars', '240', True),
    ('LH_msMincho_10p5_g280', 'ＭＳ 明朝', 21, '', 'auto', 'linesAndChars', '280', True),
    ('LH_msMincho_10p5_g312', 'ＭＳ 明朝', 21, '', 'auto', 'linesAndChars', '312', True),
    ('LH_msMincho_10p5_g360', 'ＭＳ 明朝', 21, '', 'auto', 'linesAndChars', '360', True),
    ('LH_msMincho_10p5_g400', 'ＭＳ 明朝', 21, '', 'auto', 'linesAndChars', '400', True),

    # === Vary grid_type ===
    ('LH_msMincho_10p5_gnone',  'ＭＳ 明朝', 21, '', 'auto', 'none',        '0',   True),
    ('LH_msMincho_10p5_glines', 'ＭＳ 明朝', 21, '', 'auto', 'lines',       '360', True),
    ('LH_msMincho_10p5_glAC',   'ＭＳ 明朝', 21, '', 'auto', 'linesAndChars', '360', True),

    # === Vary snap_to_grid ===
    ('LH_msMincho_10p5_g360_nosnap', 'ＭＳ 明朝', 21, '', 'auto', 'linesAndChars', '360', False),

    # === Vary explicit lineRule ===
    ('LH_msMincho_10p5_exact280', 'ＭＳ 明朝', 21, '280', 'exact',  'linesAndChars', '360', True),
    ('LH_msMincho_10p5_exact360', 'ＭＳ 明朝', 21, '360', 'exact',  'linesAndChars', '360', True),
    ('LH_msMincho_10p5_atleast', 'ＭＳ 明朝', 21, '280', 'atLeast', 'linesAndChars', '360', True),
    ('LH_msMincho_10p5_multiple_1p15', 'ＭＳ 明朝', 21, '276', 'auto', 'linesAndChars', '360', True),  # 276=23pt × 12/12.something; Multiple mapped

    # === Day 33 part 28: Times New Roman (db9ca wi=37 actual font) ===
    ('LH_TNR_10p5_g360', 'Times New Roman', 21, '', 'auto', 'linesAndChars', '360', True),
    ('LH_TNR_10p5_gnone', 'Times New Roman', 21, '', 'auto', 'none', '0', True),
    ('LH_TNR_10p5_g360_nosnap', 'Times New Roman', 21, '', 'auto', 'linesAndChars', '360', False),
    ('LH_TNR_12_g360', 'Times New Roman', 24, '', 'auto', 'linesAndChars', '360', True),
]

if __name__ == '__main__':
    for entry in TEST_MATRIX:
        build_docx(*entry)
    print(f'Wrote {len(TEST_MATRIX)} repros to {OUT}')
