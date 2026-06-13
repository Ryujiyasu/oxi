# -*- coding: utf-8 -*-
import os,sys,glob
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8')
wd=win32.gencache.EnsureDispatch('Word.Application'); wd.Visible=False
base=os.path.abspath('tools/golden-test/repros/s561_bottom_empty')
print('variant       Wpages  MARKERpage  MARKER_y  lastFiller_y')
for f in sorted(glob.glob(base+'/*.docx')):
    doc=wd.Documents.Open(f,ReadOnly=True)
    npg=doc.ComputeStatistics(2)
    mk_pg=mk_y=lf_y=None
    for i in range(1,doc.Paragraphs.Count+1):
        rng=doc.Paragraphs(i).Range; st=doc.Range(rng.Start,rng.Start)
        t=rng.Text.strip()
        if t.startswith('MARKER'):
            mk_pg=st.Information(3); mk_y=round(st.Information(6),1)
        if t.startswith('行'):
            lf_y=round(st.Information(6),1)
    print('%-22s %d      p%s       %s     %s'%(os.path.basename(f)[:-11],npg,mk_pg,mk_y,lf_y))
    doc.Close(False)
wd.Quit()
