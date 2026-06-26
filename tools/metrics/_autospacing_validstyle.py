# -*- coding: utf-8 -*-
"""Valid styled repro (Web style actually applied) + harassbosi resolved spacing."""
import os, sys, io, zipfile
sys.path.insert(0, 'tools/metrics')
import win32com.client
from mixedh_lineplace import WNS, MNS, CT, RELS, OUT
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')
FONT='ＭＳ 明朝'
def rpr(): return '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="22"/>'%(FONT,FONT,FONT)
# Complete styles.xml: docDefaults + Normal(default) + Web basedOn Normal w/ autospacing
STY=('<?xml version="1.0"?><w:styles %s>'
     '<w:docDefaults><w:rPrDefault><w:rPr>%s</w:rPr></w:rPrDefault>'
     '<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>'
     '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
     '<w:style w:type="paragraph" w:styleId="Web"><w:name w:val="Normal (Web)"/>'
     '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr>'
     '<w:rPr>%s</w:rPr></w:style></w:styles>' % (WNS, rpr(), rpr()))
def p(t,sid=None):
    ps='<w:pStyle w:val="%s"/>'%sid if sid else ''
    return '<w:p><w:pPr>%s<w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:t>%s</w:t></w:r></w:p>'%(ps,rpr(),rpr(),t)
def build(name,body):
    sect='<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/></w:sectPr>'
    doc='<?xml version="1.0"?><w:document %s %s><w:body>%s%s</w:body></w:document>'%(WNS,MNS,body,sect)
    fn=os.path.join(OUT,name)
    with zipfile.ZipFile(fn,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',CT); z.writestr('_rels/.rels',RELS)
        z.writestr('word/document.xml',doc); z.writestr('word/styles.xml',STY)
    return fn
def dump(w,fn,label):
    d=w.Documents.Open(os.path.abspath(fn),ReadOnly=True)
    print('===',label,'===')
    prev=None
    try:
        for pp in d.Paragraphs:
            y=d.Range(pp.Range.Start,pp.Range.Start).Information(6)
            gap='%.2f'%(y-prev) if prev is not None else '-'
            print('  style=%-14s Sb=%.2f Sa=%.2f y=%.2f gap=%s txt=%s'%(pp.Style.NameLocal,pp.Format.SpaceBefore,pp.Format.SpaceAfter,y,gap,pp.Range.Text.strip()[:4]))
            prev=y
    finally: d.Close(False)
def main():
    w=win32com.client.Dispatch('Word.Application'); w.Visible=False; w.DisplayAlerts=False
    try:
        # ISOLATED Web para between non-Web (the H1 vs H2 discriminator)
        dump(w, build('vs_iso.docx', p('上')+p('単独','Web')+p('下')), 'isolated Web para (H1=>0, H2=>13.75)')
        # consecutive Web
        dump(w, build('vs_run.docx', p('上')+p('W0','Web')+p('W1','Web')+p('W2','Web')+p('下')), 'consecutive Web run')
        # harassbosi resolved
        d=w.Documents.Open(os.path.abspath('tools/golden-test/documents/docx/harassbosi_002140020.docx'),ReadOnly=True)
        print('=== harassbosi first 6 Web paras: Word-resolved Sb/Sa ===')
        ps=d.Paragraphs
        for i in range(1,7):
            pp=ps(i)
            print('  style=%-14s Sb=%.2f Sa=%.2f txt=%s'%(pp.Style.NameLocal,pp.Format.SpaceBefore,pp.Format.SpaceAfter,pp.Range.Text.strip()[:6]))
        d.Close(False)
    finally: w.Quit()
if __name__=='__main__': main()
