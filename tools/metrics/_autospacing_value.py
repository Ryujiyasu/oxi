# -*- coding: utf-8 -*-
"""Pin the autospacing VALUE source: docDefaults sz dependence + in-grid behavior."""
import os, sys, io, zipfile
sys.path.insert(0, 'tools/metrics')
import win32com.client
from mixedh_lineplace import WNS, MNS, CT, RELS, OUT
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='backslashreplace')
FONT='ＭＳ 明朝'
def rpr(sz=22,extra=''): return '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/><w:sz w:val="%d"/>%s'%(FONT,FONT,FONT,sz,extra)
def para(t,sz=22,sp='',pe=''):
    s='<w:spacing %s/>'%sp if sp else ''
    return '<w:p><w:pPr>%s%s<w:rPr>%s</w:rPr></w:pPr><w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>'%(s,pe,rpr(sz),rpr(sz),t)
def build(name, body, dd_sz=None, grid=None):
    ddx = ('<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="%d"/></w:rPr></w:rPrDefault></w:docDefaults>'%dd_sz) if dd_sz else ''
    sty = '<?xml version="1.0"?><w:styles %s>%s</w:styles>'%(WNS, ddx)
    g = '<w:docGrid w:type="lines" w:linePitch="%d"/>'%grid if grid else ''
    sect='<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>%s</w:sectPr>'%g
    doc='<?xml version="1.0"?><w:document %s %s><w:body>%s%s</w:body></w:document>'%(WNS,MNS,body,sect)
    p=os.path.join(OUT,name)
    with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',CT); z.writestr('_rels/.rels',RELS)
        z.writestr('word/document.xml',doc); z.writestr('word/styles.xml',sty)
    return p
def ys(w,dx):
    d=w.Documents.Open(os.path.abspath(dx),ReadOnly=True); out=[]
    try:
        for p in d.Paragraphs:
            r=p.Range; sr=d.Range(r.Start,r.Start); out.append(sr.Information(6))
    finally: d.Close(False)
    return out
def main():
    w=win32com.client.Dispatch("Word.Application"); w.Visible=False; w.DisplayAlerts=False
    try:
        print("=== afterAuto value vs docDefaults sz (mid para fixed 11pt, nogrid) ===")
        for dd in (None,20,21,24,28):
            a=build('v_a_%s.docx'%dd, para('上',22)+para('中',22,'w:afterAutospacing="1"')+para('下',22), dd_sz=dd)
            n=build('v_n_%s.docx'%dd, para('上',22)+para('中',22)+para('下',22), dd_sz=dd)
            ya=ys(w,a); yn=ys(w,n)
            aa=(ya[2]-ya[1])-(yn[2]-yn[1])
            print("  docDefaults sz=%-5s afterAuto=%.2f  (dd_fs=%s, x1.25=%s)"%(dd, aa, dd/2 if dd else '?', round(dd/2*1.25,2) if dd else '?'))
        print("\n=== in typed grid (linePitch=360=18pt, docDefaults sz=24/12pt) ===")
        for grid in (None,360,312):
            a=build('vg_a_%s.docx'%grid, para('上',24)+para('中',24,'w:afterAutospacing="1"')+para('下',24), dd_sz=24, grid=grid)
            n=build('vg_n_%s.docx'%grid, para('上',24)+para('中',24)+para('下',24), dd_sz=24, grid=grid)
            ya=ys(w,a); yn=ys(w,n)
            aa=(ya[2]-ya[1])-(yn[2]-yn[1])
            print("  grid=%-6s gap_auto=%.2f gap_norm=%.2f afterAuto=%.2f"%(grid, ya[2]-ya[1], yn[2]-yn[1], aa))
    finally: w.Quit()
if __name__=='__main__': main()
