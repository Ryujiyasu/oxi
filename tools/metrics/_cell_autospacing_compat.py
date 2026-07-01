# -*- coding: utf-8 -*-
"""THE discriminator: does compatibilityMode decide whether before/afterAutospacing
applies (legacy HTML 13.75) or is ignored in favour of explicit before/after? (Ra 2026-07-01)

Builds docs with a settings.xml (compatibilityMode N) and measures Word's RESOLVED
Format.SpaceBefore/After for a para carrying before="100" beforeAutospacing="1"
(and variants), in BODY and in a CELL.
"""
import os, sys, io, zipfile
import win32com.client
sys.path.insert(0, 'tools/metrics')
from mixedh_lineplace import OUT, WNS, MNS
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')

FONT = 'ＭＳ 明朝'
def rpr(sz=21): return '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/>'%(FONT,FONT,FONT,sz)
def para(text, spacing): return ('<w:p><w:pPr><w:spacing %s/><w:rPr>%s</w:rPr></w:pPr>'
    '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>'%(spacing,rpr(),rpr(),text))
def pn(text): return '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:t>%s</w:t></w:r></w:p>'%(rpr(),rpr(),text)
def cell(inner): return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>'%inner)

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
 '<Default Extension="xml" ContentType="application/xml"/>'
 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
 '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
 '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')

def settings_xml(compat):
    cs = ('<w:compat><w:compatSetting w:name="compatibilityMode" '
          'w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat>'%compat) if compat else ''
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s>%s</w:settings>'%(WNS, cs)

def build(name, body, compat):
    sect=('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/></w:sectPr>')
    doc='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s %s><w:body>%s%s</w:body></w:document>'%(WNS,MNS,body,sect)
    p=os.path.join(OUT,name)
    with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',CT); z.writestr('_rels/.rels',RELS)
        z.writestr('word/_rels/document.xml.rels',DRELS)
        z.writestr('word/document.xml',doc); z.writestr('word/settings.xml',settings_xml(compat))
    return p

def measure(word, docx, target):
    doc=word.Documents.Open(os.path.abspath(docx), ReadOnly=True); res=None
    try:
        for p in doc.Paragraphs:
            if p.Range.Text.strip()==target:
                res=(p.Format.SpaceBefore, p.Format.SpaceAfter, p.Range.Information(12)); break
    finally: doc.Close(False)
    return res

# the autospacing para variants (the corpus pattern = before100+beforeAuto + after100+afterAuto)
CORPUS = 'w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"'
AUTO_ONLY = 'w:beforeAutospacing="1" w:afterAutospacing="1"'
EXPLICIT_ONLY = 'w:before="100" w:after="100"'

def main():
    word=win32com.client.Dispatch("Word.Application"); word.Visible=False; word.DisplayAlerts=False
    try:
        print("%-26s %-6s %-7s %-8s %-8s %-6s" % ("variant","compat","ctx","SpaceB","SpaceA","inTbl"))
        for label, spx in [("corpus(b100+ba+a100+aa)",CORPUS),("auto_only",AUTO_ONLY),("explicit_only(b100+a100)",EXPLICIT_ONLY)]:
            for compat in (None, 11, 14, 15):
                for ctx, body in [("body", pn('上')+para('的',spx)+pn('下')),
                                  ("cell", cell(pn('上')+para('的',spx)+pn('下')))]:
                    nm = 'cac_%s_%s_%s.docx'%(label[:6],compat,ctx)
                    r = measure(word, build(nm, body, compat), '的')
                    if r:
                        print("%-26s %-6s %-7s %-8.2f %-8.2f %-6s" % (label, str(compat), ctx, r[0], r[1], r[2]))
            print()
    finally:
        word.Quit()

if __name__=='__main__':
    main()
