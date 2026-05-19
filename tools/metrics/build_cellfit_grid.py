"""S126 cell-fit parametric repro builder.

Hypothesis (from S109e-j + 2026-05-20 visual check):
- Oxi over-fits cells: uses cell_w (full tcW) as wrap budget
- Word actually uses something smaller (inner_w = cell_w - cellMar, or less)
- For b35: 6 chars × 9.84 (compressed) = 59.04pt fits cell_w=63.55pt
  but Word fits only 4 chars at 12.45pt (expansion-fill) = 49.80pt

Goal: COM-measure Word's actual fit-width across cells of varying
(tcW, fs, content_length, jc, balanceSBDB) to derive the formula.

Output: tools/metrics/cellfit_grid/variants/cf_*.docx for COM measurement.
"""
from __future__ import annotations

import os
import zipfile
from itertools import product

OUT_DIR = os.path.join(os.path.dirname(__file__), "cellfit_grid", "variants")
os.makedirs(OUT_DIR, exist_ok=True)


def make_docx(out_path: str, tcw_dxa: int, fs_pt: float, content: str,
              jc: str, balance_sbdb: bool, compress_punct: bool):
    """Build a minimal docx with a single 1-column 1-row table containing
    a vMerge group of 4 cells (so it matches S109's 4+row vMerge trigger).

    Each cell contains `content` characters using w:rFonts MS Mincho at fs_pt.
    """
    sz_half = int(fs_pt * 2)  # half-points
    jc_tag = f'<w:jc w:val="{jc}"/>' if jc else ""
    bal_tag = "<w:balanceSingleByteDoubleByteWidth/>" if balance_sbdb else ""
    cp_tag = '<w:characterSpacingControl w:val="compressPunctuation"/>' if compress_punct else ""

    # Build 4 rows with vMerge restart on row 0, continue on 1-3
    rows = []
    for r in range(4):
        vmerge = '<w:vMerge w:val="restart"/>' if r == 0 else '<w:vMerge/>'
        cell_content = content if r == 0 else ""  # only restart row holds text
        rows.append(f"""<w:tr><w:tc><w:tcPr><w:tcW w:w="{tcw_dxa}" w:type="dxa"/>{vmerge}</w:tcPr>
<w:p><w:pPr>{jc_tag}<w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho"/><w:sz w:val="{sz_half}"/></w:rPr></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho"/><w:sz w:val="{sz_half}"/></w:rPr><w:t xml:space="preserve">{cell_content}</w:t></w:r>
</w:p>
</w:tc></w:tr>""")

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
<w:tblPr><w:tblW w:w="{tcw_dxa}" w:type="dxa"/><w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="{tcw_dxa}"/></w:tblGrid>
{''.join(rows)}
</w:tbl>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>
<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/>
</w:sectPr>
</w:body>
</w:document>
"""

    settings_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
{bal_tag}
{cp_tag}
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>
"""

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
"""

    pkg_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
"""

    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', pkg_rels)
        z.writestr('word/document.xml', document_xml)
        z.writestr('word/settings.xml', settings_xml)
        z.writestr('word/_rels/document.xml.rels', doc_rels)


# Grid: vary tcW × fs × content_length × jc × balance/compress flags
# Pin balance=on, compress=on (= b35 trigger conditions)
# Vary tcW broadly to find where the formula changes
TCW_DXAS = [1271, 1400, 1600, 1800, 2000, 2200, 2500, 3000]  # b35 col1 is 1271
FS_PT = 10.5  # b35 dominant fs
CONTENT_LENS = [4, 5, 6, 7, 8, 10, 12]
JC_VALUES = ['both', 'left', 'distribute']
BALANCE_VALUES = [True, False]

CONTENT_BASE = '組織的管理措置の見直し改善対応'  # 14 kanji chars

count = 0
for tcw, n, jc, balance in product(TCW_DXAS, CONTENT_LENS, JC_VALUES, BALANCE_VALUES):
    content = CONTENT_BASE[:n]
    name = f"cf_tcw{tcw}_n{n}_{jc}_b{int(balance)}.docx"
    out = os.path.join(OUT_DIR, name)
    make_docx(out, tcw, FS_PT, content, jc, balance, compress_punct=True)
    count += 1

print(f"built {count} variants in {OUT_DIR}")
