# -*- coding: utf-8 -*-
"""
Minimal repro to verify whether Word grid-snaps table row heights when
docGrid type="lines" is set.

Table with 5 rows, each row 1 paragraph with line=300 exact (15pt) at sz=20 (10pt).
docGrid linePitch=330 (16.5pt grid).

If Word snaps to grid: rows 16.5pt apart.
If Word doesn't snap: rows ~14-15pt apart (natural exact line height).

Measured against bd90b00 real doc shows ~14pt row gaps; need to confirm via repro.
"""
import os, zipfile
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
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho"/>
<w:sz w:val="21"/>
</w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>"""

def doc_xml(grid_type='lines', line_pitch=330, line_exact=300, sz=20, line_rule='exact', no_spacing=False):
    rows = []
    for i in range(5):
        if no_spacing:
            ppr = ''
        elif line_rule == 'auto':
            ppr = '<w:pPr></w:pPr>'  # No line spacing — defaults to single/auto
        else:
            ppr = f'<w:pPr><w:spacing w:line="{line_exact}" w:lineRule="{line_rule}"/></w:pPr>'
        rows.append(
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>'
            f'<w:p>{ppr}'
            f'<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t>行{i+1}：あいうえお</w:t></w:r></w:p></w:tc></w:tr>'
        )
    table = ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>'
             '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' +
             ''.join(rows) + '</w:tbl><w:p/>')
    sect = (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            f'w:header="720" w:footer="720" w:gutter="0"/>'
            f'<w:docGrid w:type="{grid_type}" w:linePitch="{line_pitch}"/>'
            f'</w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{table}{sect}</w:body></w:document>').encode('utf-8')

def build(name, **kwargs):
    out = OUT / name
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/document.xml', doc_xml(**kwargs))
    print(f'wrote {out}')

# RT1: replicate bd90b00 exact (line=300, pitch=330, type=lines)
build('repro_trh_RT1.docx', grid_type='lines', line_pitch=330, line_exact=300, sz=20)
# RT2: linesAndChars
build('repro_trh_RT2.docx', grid_type='linesAndChars', line_pitch=330, line_exact=300, sz=20)
# RT3: type=lines but line=240 (natural line ~14pt close to grid)
build('repro_trh_RT3.docx', grid_type='lines', line_pitch=330, line_exact=240, sz=20)
# RT4: type=lines, line=400 (above grid)
build('repro_trh_RT4.docx', grid_type='lines', line_pitch=330, line_exact=400, sz=20)

# RT5: type=lines, no line spacing (Single/auto, like real bd90b00 table cells)
build('repro_trh_RT5.docx', grid_type='lines', line_pitch=330, sz=20, line_rule='auto')
# RT6: type=linesAndChars, no line spacing
build('repro_trh_RT6.docx', grid_type='linesAndChars', line_pitch=330, sz=20, line_rule='auto')
# RT7: type=lines, no docGrid
build('repro_trh_RT7.docx', grid_type='lines', line_pitch=240, sz=20, line_rule='auto')

print('Done.')
