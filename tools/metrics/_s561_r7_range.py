# -*- coding: utf-8 -*-
"""Range-based COM (vMerge blocks Cell access): locate r7's label paragraph and
content paragraphs by Find, measure each paragraph's rendered vertical extent
(start y/page -> end y/page) to get line counts. Compare to Oxi (label ~11 @15pt,
content ~16 @14pt)."""
import sys,os
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8')
wd=win32.gencache.EnsureDispatch('Word.Application'); wd.Visible=False
path=os.path.abspath('tools/golden-test/documents/docx/roudoujoken_001161383.docx')
doc=wd.Documents.Open(path,ReadOnly=True)
USABLE=841.9-22.7-11.65; TOP=22.7
def absy(rng):
    st=doc.Range(rng.Start,rng.Start)
    return (st.Information(3)-1)*USABLE+(st.Information(6)-TOP), st.Information(3), round(st.Information(6),1)
# walk ALL paragraphs; for the form table r7 region, print each para's start abs-y, x, text
print('Word form-table paragraphs in the r7 band (abs 355..620):')
prev=None
for i in range(1,doc.Paragraphs.Count+1):
    rng=doc.Paragraphs(i).Range
    st=doc.Range(rng.Start,rng.Start)
    try:
        a=(st.Information(3)-1)*USABLE+(st.Information(6)-TOP)
    except: continue
    if 350<=a<=625:
        x=round(st.Information(5),1); pg=st.Information(3); y=round(st.Information(6),1)
        txt=rng.Text.strip()[:30]
        g=''
        if prev is not None: g='Δ%.1f'%(a-prev)
        print('  a%6.1f p%d y%6.1f x%5.1f %s  %s'%(a,pg,y,x,g,txt))
        prev=a
doc.Close(False); wd.Quit()
