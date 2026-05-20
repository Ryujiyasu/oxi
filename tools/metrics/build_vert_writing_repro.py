"""Session 130 — minimal vertical-writing repro builder.

Strips 2ea81a8441cc to a single table row that exhibits the
textDirection=tbRlV (vertical) cell + adjacent horizontal cell
pattern. This isolates the vertical writing layout question
without any other 2ea81a-specific structure.

Output: tools/golden-test/repros/vert_writing_S130/
  - V1_basic.docx       — single row, vert cell + horiz cell
  - V2_tall_text.docx   — same, but vert text is longer (overflow)
  - V3_short_text.docx  — vert text shorter than cell width
  - V4_no_valign.docx   — no vAlign (defaults to top)
  - V5_top_valign.docx  — vAlign=top explicit
"""
import os
import zipfile
import shutil

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.join(REPO_ROOT, "tools", "golden-test", "repros", "vert_writing_S130")
os.makedirs(OUT_DIR, exist_ok=True)


def make_docx(out_path: str, vert_text: str, valign: str = "center") -> None:
    """Write a minimal docx with one 2-cell table row.

    Cell 1: 709dxa wide, textDirection=tbRlV, vAlign=valign, contains vert_text.
    Cell 2: 8647dxa wide, normal, contains 4 horizontal paragraphs.
    """
    ct = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

    word_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho" w:hAnsi="MS Mincho" w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>'''

    # vAlign element (None if explicitly skipped)
    valign_xml = f'<w:vAlign w:val="{valign}"/>' if valign else ''

    # Build a heading paragraph (matches 2ea81a structure with sz=16 in vertical cell)
    vert_text_xml = f'<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/><w:rPr><w:sz w:val="16"/><w:szCs w:val="21"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="16"/><w:szCs w:val="21"/></w:rPr><w:t>{vert_text}</w:t></w:r></w:p>'

    # Build the horizontal cell with 4 paragraphs
    horiz_paras = []
    horiz_paras.append('<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="20"/></w:rPr><w:t>（いずれかを選択）</w:t></w:r></w:p>')
    horiz_paras.append('<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="20"/></w:rPr><w:t>１．修正申告書等をおおむね６月以内に提出予定</w:t></w:r></w:p>')
    horiz_paras.append('<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="20"/></w:rPr><w:t>２．期限内申告書をおおむね12月以内に提出予定</w:t></w:r></w:p>')
    horiz_paras.append('<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="20"/></w:rPr><w:t>３．その他（具体的な理由を記載してください。）</w:t></w:r></w:p>')
    horiz_cell_content = ''.join(horiz_paras)

    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:body>
<w:p><w:r><w:t>Header paragraph (anchor reference)</w:t></w:r></w:p>
<w:tbl>
<w:tblPr><w:tblW w:w="9356" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="709"/><w:gridCol w:w="8647"/></w:tblGrid>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="709" w:type="dxa"/><w:textDirection w:val="tbRlV"/>{valign_xml}</w:tcPr>{vert_text_xml}</w:tc>
<w:tc><w:tcPr><w:tcW w:w="8647" w:type="dxa"/></w:tcPr>{horiz_cell_content}</w:tc>
</w:tr>
</w:tbl>
<w:p><w:r><w:t>Footer paragraph (anchor reference)</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>
</w:body>
</w:document>'''

    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/_rels/document.xml.rels', word_rels)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/document.xml', document)
    print(f'  Wrote: {out_path}')


def main():
    variants = [
        ('V1_basic.docx',      '予納する理由',   'center'),
        ('V2_tall_text.docx',  '予納する理由を選択する場合',  'center'),  # ~13 chars
        ('V3_short_text.docx', '理由',           'center'),
        ('V4_no_valign.docx',  '予納する理由',   None),       # no vAlign
        ('V5_top_valign.docx', '予納する理由',   'top'),
    ]
    print(f'Generating {len(variants)} repros in {OUT_DIR}:')
    for fn, vert_text, valign in variants:
        out = os.path.join(OUT_DIR, fn)
        make_docx(out, vert_text, valign)
    print('Done.')


if __name__ == '__main__':
    main()
