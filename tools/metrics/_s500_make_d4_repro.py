# -*- coding: utf-8 -*-
"""L2 repro: match d4d126's top-row structure to isolate the +3.3 too-low cell-Y.
docGrid linesAndChars linePitch=292; settings adjustLineHeightInTable; a 2-col row
trHeight=658 (atLeast, no hRule), vAlign=center; cell0 first para lineRule=exact line=240
+ beforeLines toggle, sz=20 (10pt) MS Mincho; cell1 a normal tall-ish cell. Build two
variants: with/without beforeLines, to isolate whether space-before is wrongly in the
vAlign=center content height. cp932-safe (CJK as literals in UTF-8 file), raw OOXML.
"""
import os, zipfile

OUT = os.path.join(os.path.dirname(__file__), '..', 'golden-test', 'repros', 'd4_cellY')
os.makedirs(OUT, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
A = 'あ'  # あ

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
WRELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''
SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings %s><w:compat><w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>''' % NS


def cell(content, valign='center'):
    return ('<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vAlign w:val="%s"/></w:tcPr>%s</w:tc>'
            % (valign, content))


def para(text, before_lines):
    sb = ('<w:spacing w:beforeLines="30" w:before="87" w:line="240" w:lineRule="exact"/>'
          if before_lines else '<w:spacing w:line="240" w:lineRule="exact"/>')
    rpr = '<w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝" w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
    return '<w:p><w:pPr>%s%s</w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>' % (sb, rpr, rpr, text)


def tall_cell_paras(n):
    out = ''
    for i in range(n):
        out += para(A * 3, False)
    return out


def doc_xml(before_lines):
    # cell0 short (1 char) vAlign=center exact-line +/- beforeLines; cell1 tall (3 exact lines)
    c0 = cell(para(A, before_lines))
    c1 = cell(tall_cell_paras(3))
    row = '<w:tr><w:trPr><w:trHeight w:val="658"/></w:trPr>%s%s</w:tr>' % (c0, c1)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>%s</w:tbl>' % row)
    ref = '<w:p><w:r><w:t>REF</w:t></w:r></w:p>'
    after = '<w:p><w:r><w:t>AFTER</w:t></w:r></w:p>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
            '<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s%s%s</w:body></w:document>'
            % (NS, ref, tbl, after, sect))


def build(name, before_lines):
    path = os.path.join(OUT, name)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', WRELS)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', doc_xml(before_lines))
    return path


print('built', build('d4_sb.docx', True))
print('built', build('d4_nosb.docx', False))
print('DONE')
