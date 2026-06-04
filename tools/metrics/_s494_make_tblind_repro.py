# -*- coding: utf-8 -*-
"""Build a minimal repro docx isolating the autofit-table tblInd + cell-margin hypothesis.
Two single-cell tables, identical except tblW: one auto (autofit), one fixed (dxa). Each
has tblInd=817, cellMar left=108. If Oxi over-indents the AUTOFIT cell text by 108tw (5.4pt)
vs Word but the FIXED one matches, the hypothesis (Oxi double-counts cell margin on autofit
tblInd) is confirmed. cp932-safe (CJK text is built from code points, not literals)."""
import zipfile, os

# あいうえお as code points (avoid cp932 source-literal mangling)
CJK = ''.join(chr(c) for c in [0x3042, 0x3044, 0x3046, 0x3048, 0x304A])

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''


def tbl(tblw_type, tblw_val, label):
    return f'''<w:tbl>
<w:tblPr>
<w:tblW w:w="{tblw_val}" w:type="{tblw_type}"/>
<w:tblInd w:w="817" w:type="dxa"/>
<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>
<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>
</w:tblPr>
<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
<w:tr><w:tc>
<w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
<w:p><w:pPr><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho"/><w:sz w:val="21"/></w:rPr></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho"/><w:sz w:val="21"/></w:rPr><w:t>{label}{CJK}</w:t></w:r></w:p>
</w:tc></w:tr>
</w:tbl>
<w:p/>'''


DOC = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
{tbl("auto", "0", "A")}
{tbl("dxa", "4000", "F")}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304" w:header="851" w:footer="992" w:gutter="0"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>'''

out = 'tools/golden-test/repros/tblind_autofit/tblind_repro.docx'
os.makedirs(os.path.dirname(out), exist_ok=True)
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT)
    z.writestr('_rels/.rels', RELS)
    z.writestr('word/document.xml', DOC)
print('wrote', out)
print('margin=65.2pt, tblInd=817tw=40.85pt -> Word cell text x expected ~106.0pt (margin+tblInd)')
print('if Oxi autofit cell at ~111.4 (=+108tw cellMar) but fixed at 106.0 -> hypothesis CONFIRMED')
