# -*- coding: utf-8 -*-
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
DOCX=os.path.abspath('tools/golden-test/documents/docx/kojin_000505813.docx')
w=win32.DispatchEx('Word.Application'); w.Visible=False
try:
    d=w.Documents.Open(DOCX,ReadOnly=True)
    target=None
    for p in d.Paragraphs:
        if '以下の各方法により' in p.Range.Text:
            target=p; break
    if target is None: print('not found'); 
    else:
        rng=target.Range; n=min(rng.Characters.Count,40)
        print('font.Name=%s size=%s'%(rng.Font.Name, rng.Font.Size))
        prev=None
        for k in range(1,n+1):
            c=rng.Characters(k); x=c.Information(5)
            adv=round(x-prev,2) if prev is not None else 0
            print('  %-2s adv=%5.2f'%(c.Text.replace(chr(13),'CR')[:2],adv)); prev=x
    d.Close(False)
finally:
    w.Quit()
