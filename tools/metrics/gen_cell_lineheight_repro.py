# -*- coding: utf-8 -*-
"""S492v — generate minimal cell-line-height repro docx files matching b35's docGrid.
One single-cell table per file; cell holds CJK text long enough to wrap to several lines;
snapToGrid=0 (matching b35's 13 opt-out paras); docGrid linesAndChars linePitch=350
charSpace=-2714. Vary eastAsia font x size. COM then measures the intra-cell line pitch
(pure line height, no para spacing) to derive Word's per-font cell-line-height function.
cp932-safe: this file is UTF-8 (authored via Write), zip writes bytes. Output: c:/tmp/cellrepro/."""
import os, zipfile

OUT = r'c:\tmp\cellrepro'
os.makedirs(OUT, exist_ok=True)

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

DOCRELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>
</w:styles>'''

# CJK text long enough to wrap to several lines in a ~4000tw cell
TEXT = '東京都の電子計算機処理等に係る個人情報の保護に関する管理規程について定めるものとする。'


def make_doc(ea_font, sz_half):
    # sz_half = half-points (e.g. 21 = 10.5pt). snapToGrid=0, line spacing default (single).
    rpr = ('<w:rPr><w:rFonts w:ascii="Century" w:eastAsia="%s" w:hAnsi="Century"/>'
           '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (ea_font, sz_half, sz_half))
    # two paragraphs in the cell, each wrapping, so we capture many intra-cell lines
    para = ('<w:p><w:pPr><w:snapToGrid w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="both"/>%s</w:pPr>'
            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (rpr, rpr, TEXT * 2))
    cell = ('<w:tc><w:tcPr><w:tcW w:w="4200" w:type="dxa"/></w:tcPr>%s</w:tc>' % para)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="4200" w:type="dxa"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="4200"/></w:tblGrid>'
           '<w:tr>%s</w:tr></w:tbl>' % cell)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s<w:p/>'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
           '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/></w:sectPr>'
           '</w:body></w:document>' % tbl)
    return doc


FONTS = [('ＭＳ 明朝', 'MSMincho'), ('ＭＳ ゴシック', 'MSGothic')]
SIZES = [(18, '9'), (20, '10'), (21, '10p5'), (22, '11'), (24, '12')]

made = []
for ea, ftag in FONTS:
    for szh, stag in SIZES:
        name = 'cell_%s_%s.docx' % (ftag, stag)
        path = os.path.join(OUT, name)
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/document.xml', make_doc(ea, szh))
        made.append((name, ea, szh / 2.0))

with open(os.path.join(OUT, 'manifest.txt'), 'w', encoding='utf-8') as f:
    for name, ea, sz in made:
        f.write('%s\t%s\t%.1f\n' % (name, ea, sz))
print('generated %d repro docx in %s' % (len(made), OUT))
for name, ea, sz in made:
    print('  %s  (%.1fpt)' % (name, sz))
