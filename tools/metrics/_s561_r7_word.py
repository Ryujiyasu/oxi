# -*- coding: utf-8 -*-
"""Word COM: form-table row r7 (Rows(8)) cell content extents + line counts, to
pin whether the 12.8pt Oxi under-count is line-COUNT or line-HEIGHT, and which
cell dominates."""
import sys,os
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8')
wd=win32.gencache.EnsureDispatch('Word.Application'); wd.Visible=False
path=os.path.abspath('tools/golden-test/documents/docx/roudoujoken_001161383.docx')
doc=wd.Documents.Open(path,ReadOnly=True)
tbl=doc.Tables(1)
print('table rows=%d cols(r8)=%d'%(tbl.Rows.Count, tbl.Rows(8).Cells.Count))
r=tbl.Rows(8)
for ci in range(1,r.Cells.Count+1):
    c=r.Cells(ci)
    rng=c.Range
    # collapse start
    st=doc.Range(rng.Start,rng.Start)
    en=doc.Range(rng.End-1,rng.End-1)
    y0=st.Information(6); p0=st.Information(3)
    y1=en.Information(6); p1=en.Information(3)
    txt=rng.Text.strip()[:24]
    # width
    try: w=c.Width
    except: w=None
    # count lines: walk the cell range line by line via Information(6) of each char? approx via paras
    npara=rng.Paragraphs.Count
    print('  cell%d w=%.1f npara=%d  start(p%d y%.1f) end(p%d y%.1f)  %r'%(ci,w or -1,npara,p0,y0,p1,y1,txt))
doc.Close(False); wd.Quit()
