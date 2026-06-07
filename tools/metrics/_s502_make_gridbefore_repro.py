# -*- coding: utf-8 -*-
"""S502 gridBefore repro (the angle S494b never built): isolate the +4.1pt over-indent on
mode-15 gridBefore tables (29dc6e/15076). A 3-col table, tblInd present, with rows of
gridBefore 0 / 1 / 2 (skipping leading grid columns). Measure where Word places the leading
cell content of each row vs Oxi. If gridBefore rows diverge but gridBefore=0 matches, the
skipped-column-width (or tblInd-with-gridBefore) is the mechanism. cp932-safe (CJK literals
in UTF-8 file), raw OOXML. Matches 29dc6e: tblInd~420, cellMar=12, docGrid linesAndChars."""
import os, zipfile
OUT = os.path.join(os.path.dirname(__file__), '..', 'golden-test', 'repros', 'gridbefore')
os.makedirs(OUT, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
A = 'あ'

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
WRELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''
SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings %s><w:compat><w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>''' % NS

RPR = '<w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝" w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'

def cell(txt, w):
    body = ('<w:p><w:pPr>%s</w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>' % (RPR, RPR, txt)) if txt else '<w:p/>'
    return '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc>' % (w, body)

def row(gb, label):
    # gb = gridBefore count; the leading visible cell carries `label` (a digit to find x)
    gbpr = ('<w:gridBefore w:val="%d"/>' % gb) if gb else ''
    cells = ''
    # remaining visible cells = 3 - gb
    widths = [1800, 1800, 1800]
    for i in range(gb, 3):
        cells += cell(label if i == gb else A, widths[i])
    return '<w:tr><w:trPr>%s</w:trPr>%s</w:tr>' % (gbpr, cells)

def doc_xml():
    rows = row(0, '0') + row(1, '1') + row(2, '2')  # gridBefore 0,1,2; leading cell labeled
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="420" w:type="dxa"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           '<w:tblCellMar><w:left w:w="12" w:type="dxa"/><w:right w:w="12" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid>%s</w:tbl>' % rows)
    ref = '<w:p><w:r><w:t>REF</w:t></w:r></w:p>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
            '<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/></w:sectPr>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s%s</w:body></w:document>' % (NS, ref, tbl, sect)

path = os.path.join(OUT, 'gb.docx')
with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT)
    z.writestr('_rels/.rels', RELS)
    z.writestr('word/_rels/document.xml.rels', WRELS)
    z.writestr('word/settings.xml', SETTINGS)
    z.writestr('word/document.xml', doc_xml())
print('built', path)
