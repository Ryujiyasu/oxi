# -*- coding: utf-8 -*-
import sys, os, json
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8')
wd = win32.gencache.EnsureDispatch('Word.Application'); wd.Visible=False
for name in ('s560short_repro.docx','s560tall_repro.docx'):
    path=os.path.abspath('tools/golden-test/repros/s560_cont_cols/'+name)
    doc=wd.Documents.Open(path,ReadOnly=True)
    print('=== %s : %d sections ==='%(name,doc.Sections.Count))
    for s in range(1,doc.Sections.Count+1):
        sec=doc.Sections(s)
        tc=sec.PageSetup.TextColumns
        first=sec.Range.Paragraphs(1).Range.Text.strip()[:12]
        print('  sec%d  TextColumns.Count=%d  firstpara=%s'%(s,tc.Count,first))
    doc.Close(False)
wd.Quit()
