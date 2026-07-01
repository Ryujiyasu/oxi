# -*- coding: utf-8 -*-
"""Discriminator via RENDERED Y-GAPS (Format.SpaceBefore is unreliable — it returns
only explicit settings, 0 for auto_only). Measures the ACTUAL rendered before/after
space for the corpus pattern (before=100 + beforeAutospacing) across compatibilityMode,
in body and cell. (Ra 2026-07-01)
"""
import os, sys, io, zipfile
import win32com.client
sys.path.insert(0, 'tools/metrics')
from mixedh_lineplace import OUT, WNS, MNS
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')

FONT='ＭＳ 明朝'
def rpr(sz=21): return '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/>'%(FONT,FONT,FONT,sz)
def para(text,spx): return '<w:p><w:pPr><w:spacing %s/><w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>'%(spx,rpr(),rpr(),text)
def pn(text): return '<w:p><w:pPr><w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:t>%s</w:t></w:r></w:p>'%(rpr(),rpr(),text)
def cell(inner): return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
 '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
 '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr>'
 '<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>'%inner)
CT=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
 '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
def settings_xml(c):
    cs=('<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat>'%c) if c else ''
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s>%s</w:settings>'%(WNS,cs)
def build(name,body,compat):
    sect='<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/></w:sectPr>'
    doc='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s %s><w:body>%s%s</w:body></w:document>'%(WNS,MNS,body,sect)
    p=os.path.join(OUT,name)
    with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',CT); z.writestr('_rels/.rels',RELS)
        z.writestr('word/_rels/document.xml.rels',DRELS); z.writestr('word/document.xml',doc)
        z.writestr('word/settings.xml',settings_xml(compat))
    return p
def ys(word,docx):
    doc=word.Documents.Open(os.path.abspath(docx),ReadOnly=True); out=[]
    try:
        for p in doc.Paragraphs:
            sr=doc.Range(p.Range.Start,p.Range.Start); out.append((sr.Information(6),p.Range.Text.strip()))
    finally: doc.Close(False)
    return out

# patterns: each renders TOP / MID(spacing) / BOT ; compare MID gaps vs a plain control
PATTERNS = {
 'corpus_b100ba_a100aa': 'w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"',
 'auto_only_baaa'      : 'w:beforeAutospacing="1" w:afterAutospacing="1"',
 'explicit_b100a100'   : 'w:before="100" w:after="100"',
}
def main():
    word=win32com.client.Dispatch("Word.Application"); word.Visible=False; word.DisplayAlerts=False
    try:
        print("%-22s %-6s %-6s %-8s %-8s %-8s %-8s" % ("pattern","compat","ctx","beforeGap","afterGap","befSpace","aftSpace"))
        for pname, spx in PATTERNS.items():
            for compat in (None, 11, 14, 15):
                for ctx in ('body','cell'):
                    mk = (lambda b: cell(b)) if ctx=='cell' else (lambda b: b)
                    auto = ys(word, build('cc2_%s_%s_%s.docx'%(pname[:8],compat,ctx),
                        mk(pn('上')+para('的',spx)+pn('下')), compat))
                    norm = ys(word, build('cc2_n_%s_%s.docx'%(compat,ctx),
                        mk(pn('上')+pn('的')+pn('下')), compat))
                    bg=auto[1][0]-auto[0][0]; ag=auto[2][0]-auto[1][0]
                    bn=norm[1][0]-norm[0][0]; an=norm[2][0]-norm[1][0]
                    print("%-22s %-6s %-6s %-8.2f %-8.2f %-8.2f %-8.2f" % (pname,str(compat),ctx,bg,ag,bg-bn,ag-an))
            print()
    finally:
        word.Quit()
if __name__=='__main__': main()
