# -*- coding: utf-8 -*-
"""S502 cell-content-position repro: replicate 29dc6e p5's content cell to isolate the
uniform +4.1 cell-content-left shift (advance matches, columns match per Word border tool).
A 2-col table (col0 ~105pt so col1 starts ~x160), col1 carries a content para. Variants
toggle jc (center/left) and firstLine (on/off) under docGrid linesAndChars (linePitch=292,
charSpace=1453), sz=24 (12pt) MS Mincho. Measure col1 first-char x vs Word -> isolate whether
the +4.1 is jc=center, firstLine, or char-grid phase. cp932-safe, raw OOXML."""
import os, zipfile
OUT = os.path.join(os.path.dirname(__file__), '..', 'golden-test', 'repros', 'cellpos')
os.makedirs(OUT, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CJK = 'あいう'  # SHORT: must NOT wrap so line_total_w is identical across variants (isolates indent/center)

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
WRELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''
SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings %s><w:compat><w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>''' % NS
RPR = '<w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝" w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'

def content_para(jc, firstline):
    ind = '<w:ind w:firstLineChars="100" w:firstLine="247"/>' if firstline else ''
    jct = '<w:jc w:val="%s"/>' % jc if jc else ''
    return '<w:p><w:pPr>%s%s</w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>' % (jct, ind, RPR, CJK)

def doc_xml(jc, firstline):
    c0 = '<w:tc><w:tcPr><w:tcW w:w="2110" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
    c1 = '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>%s</w:tc>' % content_para(jc, firstline)
    row = '<w:tr>%s%s</w:tr>' % (c0, c1)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           '<w:tblCellMar><w:left w:w="12" w:type="dxa"/><w:right w:w="12" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="2110"/><w:gridCol w:w="6000"/></w:tblGrid>%s</w:tbl>' % row)
    ref = '<w:p><w:r><w:t>REF</w:t></w:r></w:p>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
            '<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/></w:sectPr>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s%s</w:body></w:document>' % (NS, ref, tbl, sect)

def build(name, jc, firstline):
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', WRELS); z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', doc_xml(jc, firstline))
    return p

print('built', build('cp_center_fl.docx', 'center', True))
print('built', build('cp_left_fl.docx', 'left', True))
print('built', build('cp_center_nofl.docx', 'center', False))
print('built', build('cp_left_nofl.docx', 'left', False))
print('DONE')
