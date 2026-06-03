# -*- coding: utf-8 -*-
"""S492w — controlled repro to CONFIRM Word's cell-content line-height RULE under b35's
exact conditions. Adds <w:adjustLineHeightInTable/> (b35 HAS it; the S492v repro did NOT,
which is why it gave natural 14.0) + b35's exact docGrid (linesAndChars linePitch=350
charSpace=-2714). Two paragraph variants per size:
  - default snapToGrid (true)  -> matches b35's 10.5pt cells (Word: 15.0/17.2 spread?)
  - snapToGrid=0               -> matches b35's 9pt opt-out cells (Word: natural 12.0)
Vary font size 9/10/10.5/11/12. Long wrapping CJK text -> multi-line, so COM can measure the
intra-cell line pitch. Determines whether Word: snaps to linePitch (17.5), uses sub-grid
(15.0), or natural (14.0). cp932-safe (UTF-8 file). Output: c:/tmp/cellflag/."""
import os, zipfile

OUT = r'c:\tmp\cellflag'
os.makedirs(OUT, exist_ok=True)

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOCRELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>
</w:styles>'''

# THE KEY: adjustLineHeightInTable present (b35 has it)
SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:adjustLineHeightInTable/>
</w:settings>'''

TEXT = '東京都の電子計算機処理等に係る個人情報の保護に関する管理規程について定めるものとする。'


def make_doc(sz_half, snap0):
    sg = '<w:snapToGrid w:val="0"/>' if snap0 else ''
    rpr = ('<w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
           '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (sz_half, sz_half))
    para = ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="both"/>%s</w:pPr>'
            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (sg, rpr, rpr, TEXT * 2))
    cell = ('<w:tc><w:tcPr><w:tcW w:w="4200" w:type="dxa"/></w:tcPr>%s</w:tc>' % para)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="4200" w:type="dxa"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
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


SIZES = [(18, '9'), (20, '10'), (21, '10p5'), (22, '11'), (24, '12')]
made = []
for szh, stag in SIZES:
    for snap0, sgtag in [(False, 'snapON'), (True, 'snap0')]:
        name = 'flag_%s_%s.docx' % (stag, sgtag)
        path = os.path.join(OUT, name)
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT)
            z.writestr('_rels/.rels', RELS)
            z.writestr('word/_rels/document.xml.rels', DOCRELS)
            z.writestr('word/styles.xml', STYLES)
            z.writestr('word/settings.xml', SETTINGS)
            z.writestr('word/document.xml', make_doc(szh, snap0))
        made.append((name, szh / 2.0, snap0))

with open(os.path.join(OUT, 'manifest.txt'), 'w', encoding='utf-8') as f:
    for name, sz, s0 in made:
        f.write('%s\t%.1f\t%d\n' % (name, sz, 1 if s0 else 0))
print('generated %d flagged repro docx in %s (adjustLineHeightInTable=ON, docGrid linesAndChars 350)' % (len(made), OUT))
for name, sz, s0 in made:
    print('  %s  (%.1fpt, snapToGrid=%s)' % (name, sz, '0' if s0 else 'default'))
