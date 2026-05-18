"""S107 minimal repro: measure FIRST LINE Y position for grid-snapped Single
paragraphs across various fonts and sizes.

Goal: discover Word's natural_lh formula for grid-snap half-leading.
"""
import os, zipfile
from pathlib import Path

OUT = Path('c:/Users/ryuji/oxi-main/tools/metrics/first_line_offset_repro')
OUT.mkdir(parents=True, exist_ok=True)

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:kern w:val="2"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a">
<w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>
</w:style>
</w:styles>'''


def make_docx(name, font, font_sz_hps, line_pitch=360):
    """font_sz_hps = font size in half-points (sz=24 means 12pt)."""
    rfonts = f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:hint="eastAsia"/>'
    paras = []
    for txt in ['Line1', 'Line2', 'Line3']:
        rpr = f'<w:rPr>{rfonts}<w:kern w:val="2"/><w:sz w:val="{font_sz_hps}"/></w:rPr>'
        paras.append(f'<w:p><w:pPr><w:rPr>{rfonts}<w:sz w:val="{font_sz_hps}"/></w:rPr></w:pPr><w:r>{rpr}<w:t>{txt}</w:t></w:r></w:p>')
    body = '\n'.join(paras) + f'''
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397"/>
<w:docGrid w:type="lines" w:linePitch="{line_pitch}"/>
</w:sectPr>'''
    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>{body}</w:body></w:document>'''
    out_path = OUT / f"{name}.docx"
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc_xml)


def main():
    # Various font/size combinations, all grid linePitch=360
    cases = [
        ('Gothic_10_5', 'ＭＳ ゴシック', 21),  # 10.5pt
        ('Gothic_12', 'ＭＳ ゴシック', 24),    # 12pt
        ('Gothic_14', 'ＭＳ ゴシック', 28),    # 14pt
        ('Gothic_16', 'ＭＳ ゴシック', 32),    # 16pt
        ('Mincho_10_5', 'ＭＳ 明朝', 21),
        ('Mincho_12', 'ＭＳ 明朝', 24),
        ('Mincho_14', 'ＭＳ 明朝', 28),
        ('TNR_10_5', 'Times New Roman', 21),
        ('TNR_12', 'Times New Roman', 24),
        ('Calibri_11', 'Calibri', 22),
    ]
    for name, font, sz in cases:
        make_docx(name, font, sz, line_pitch=360)
    # Different grid pitches
    make_docx('Gothic_12_pitch300', 'ＭＳ ゴシック', 24, line_pitch=300)
    make_docx('Gothic_12_pitch400', 'ＭＳ ゴシック', 24, line_pitch=400)
    print(f"Built {len(cases) + 2} repros")


if __name__ == '__main__':
    main()
