# -*- coding: utf-8 -*-
"""S560: COM-measure the continuous-column-change repros. Per paragraph:
page, x (Information(5) horiz pos rel page), y (Information(6) vert pos rel
page). Uses collapsed start range (R30 fix)."""
import sys, os
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8')
wd = win32.gencache.EnsureDispatch('Word.Application')
wd.Visible = False
base = os.path.abspath('tools/golden-test/repros/s560_cont_cols')
for name in ('s560short_repro.docx','s560tall_repro.docx'):
    path = os.path.join(base, name)
    doc = wd.Documents.Open(path, ReadOnly=True)
    print('=== %s : pages=%d ==='%(name, doc.ComputeStatistics(2)))  # 2=wdStatisticPages
    for i in range(1, doc.Paragraphs.Count+1):
        rng = doc.Paragraphs(i).Range
        start = doc.Range(rng.Start, rng.Start)
        pg = start.Information(3)   # wdActiveEndPageNumber
        x  = start.Information(5)   # wdHorizontalPositionRelativeToPage
        y  = start.Information(6)   # wdVerticalPositionRelativeToPage
        txt = rng.Text.strip()[:14]
        print('  P%-2d p%d x=%6.1f y=%6.1f  %s'%(i,pg,x,y,txt))
    doc.Close(False)
wd.Quit()
