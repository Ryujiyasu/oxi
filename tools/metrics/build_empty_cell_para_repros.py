"""Build minimal repro docx fixtures for empty-cell-paragraph height behavior.

R49 (2026-04-29): R48 instrumentation found Oxi over-estimates empty cell
paragraph heights. This script generates 6 fixture variants, each isolating
one parameter of the empty-paragraph rendering rule, so we can COM-measure
Word's actual height and derive the correct rule.

Each fixture is a 1-cell table with:
  - 1 empty paragraph (the variant under test)
  - 1 marker paragraph (so we can measure cell height as marker_y - empty_y)

Output: tests/fixtures/empty_cell_para/{V1..V6}.docx
"""
from __future__ import annotations
import os, zipfile, datetime

OUT = os.path.join(os.path.dirname(__file__), '..', '..', 'tests', 'fixtures', 'empty_cell_para')

# Body XML stub: a 1-row table with two cells. First cell holds variant
# empty paragraph + marker. Second cell holds a single short paragraph
# (so we can compare cell heights).
BODY_TEMPLATE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>HEADER paragraph</w:t></w:r></w:p>
<w:tbl>
<w:tblPr>
  <w:tblW w:w="9000" w:type="dxa"/>
  <w:tblBorders>
    <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  </w:tblBorders>
</w:tblPr>
<w:tblGrid>
  <w:gridCol w:w="4500"/>
  <w:gridCol w:w="4500"/>
</w:tblGrid>
<w:tr>
  <w:tc>
    <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
    {VARIANT_PARA}
    <w:p><w:r><w:t>marker</w:t></w:r></w:p>
  </w:tc>
  <w:tc>
    <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
    <w:p><w:r><w:t>right</w:t></w:r></w:p>
  </w:tc>
</w:tr>
</w:tbl>
<w:p><w:r><w:t>FOOTER paragraph</w:t></w:r></w:p>
<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
  <w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>'''

VARIANTS = {
    'V1_bare':           '<w:p/>',
    'V2_pPr_only':       '<w:p><w:pPr/></w:p>',
    'V3_rPr_sz10':       '<w:p><w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr></w:p>',
    'V4_rPr_sz14':       '<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr></w:p>',
    'V5_line_exact_300': '<w:p><w:pPr><w:spacing w:line="300" w:lineRule="exact"/></w:pPr></w:p>',
    'V6_line_auto_240':  '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:p>',
}

# Minimal supporting files for valid docx
CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''


def build(name: str, variant_xml: str) -> None:
    body = BODY_TEMPLATE.replace('{VARIANT_PARA}', variant_xml)
    out_path = os.path.join(OUT, f'{name}.docx')
    os.makedirs(OUT, exist_ok=True)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/document.xml', body)
    print(f'  built {out_path}')


def main() -> None:
    print(f'Output dir: {os.path.abspath(OUT)}')
    for name, xml in VARIANTS.items():
        build(name, xml)
    print(f'\\n{len(VARIANTS)} fixtures built.')


if __name__ == '__main__':
    main()
