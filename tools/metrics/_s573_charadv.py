# -*- coding: utf-8 -*-
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
DOCX = os.path.abspath('tools/golden-test/documents/docx/ikujikaigo_001676343.docx')
w = win32.DispatchEx('Word.Application'); w.Visible=False
try:
    d = w.Documents.Open(DOCX, ReadOnly=True)
    for ti in (int(sys.argv[1]) if len(sys.argv)>1 else 41,):
        p=d.Paragraphs(ti); rng=p.Range
        n=rng.Characters.Count
        prev=None; out=[]
        for k in range(1,n+1):
            c=rng.Characters(k); x=c.Information(5); y=c.Information(6)
            out.append((c.Text,x,y)); 
        print('=== i=%d chars=%d ==='%(ti,n))
        prevy=None
        for ch,x,y in out:
            nl = (prevy is not None and abs(y-prevy)>1.0)
            disp=ch.replace('\r','CR').replace('　','　')
            adv = (x-prev) if (prev is not None and not nl) else 0
            mark=' <<<NEWLINE y=%.1f'%y if nl else ''
            print('  %-2s x=%6.2f adv=%5.2f%s'%(disp[:2],x,adv,mark))
            prev=x; prevy=y
    d.Close(False)
finally:
    w.Quit()
