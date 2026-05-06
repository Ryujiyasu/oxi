# -*- coding: utf-8 -*-
"""
Generate a minimal repro doc to test footer-distance reservation when no footer.xml exists.

Variants written to tools/golden-test/documents/docx/repro_footer_floor_*.docx :

  V1: bottom=100tw (5pt),   footer=1440tw (72pt)  — large fd reservation expected
  V2: bottom=2000tw (100pt), footer=1440tw (72pt) — bm > fd, should ignore fd
  V3: bottom=500tw  (25pt),  footer=720tw (36pt)  — moderate gap

Body: 80 paragraphs of plain CJK text "あいうえお" with line=240 exact (12pt) at 12pt MS Mincho.
That gives 12pt per line — predictable cumulative y. We can then COM-measure each page's
max_body_y and verify against (pgH - max(bm, fd)).
"""
import os, sys, zipfile
from pathlib import Path

OUT = Path('tools/metrics/_repros')
OUT.mkdir(parents=True, exist_ok=True)

CT = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr></w:pPr></w:pPrDefault>
</w:docDefaults>
</w:styles>"""

def doc_xml(bot_tw, footer_tw, n_paras=80, title_pg=False, doc_grid_type='default', line_pitch=312):
    paras = []
    for i in range(n_paras):
        paras.append(
            f'<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho"/><w:sz w:val="24"/></w:rPr>'
            f'<w:t xml:space="preserve">L{i:02d} あいうえおかきくけこ</w:t></w:r></w:p>'
        )
    body = ''.join(paras)
    title_pg_xml = '<w:titlePg/>' if title_pg else ''
    sect = (
        f'<w:sectPr>'
        f'<w:pgSz w:w="11906" w:h="16838"/>'
        f'<w:pgMar w:top="1440" w:right="1440" w:bottom="{bot_tw}" w:left="1440" '
        f'w:header="720" w:footer="{footer_tw}" w:gutter="0"/>'
        f'<w:cols w:space="425"/>'
        f'{title_pg_xml}'
        f'<w:docGrid w:type="{doc_grid_type}" w:linePitch="{line_pitch}"/>'
        f'</w:sectPr>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body}{sect}</w:body></w:document>'
    ).encode('utf-8')

def build(out_path, bot_tw, footer_tw, **kwargs):
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/document.xml', doc_xml(bot_tw, footer_tw, **kwargs))

variants = [
    # name, bot_tw, footer_tw, kwargs
    ('repro_footer_floor_V1.docx', 100,  1440, {}),                                    # baseline (already tested)
    ('repro_footer_floor_V2.docx', 2000, 1440, {}),
    ('repro_footer_floor_V3.docx', 500,  720,  {}),
    # New: 2ea81a-like config (titlePg + docGrid type=lines + linePitch=323)
    ('repro_footer_floor_V4.docx', 397,  907,  dict(title_pg=True, doc_grid_type='lines', line_pitch=323)),
    # Isolate which attribute matters
    ('repro_footer_floor_V5.docx', 397,  907,  dict(title_pg=False, doc_grid_type='lines', line_pitch=323)),  # docGrid=lines only
    ('repro_footer_floor_V6.docx', 397,  907,  dict(title_pg=True, doc_grid_type='default', line_pitch=312)), # titlePg only
]

for name, bot, footer, kw in variants:
    build(OUT / name, bot, footer, **kw)
    print(f'wrote {OUT/name}  bot={bot}tw  footer={footer}tw  kw={kw}')

print('Done.')
