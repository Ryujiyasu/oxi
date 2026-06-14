# -*- coding: utf-8 -*-
import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
DOCX=os.path.abspath('tools/golden-test/documents/docx/ikujikaigo_001676343.docx')
w=win32.DispatchEx('Word.Application'); w.Visible=False
try:
    d=w.Documents.Open(DOCX,ReadOnly=True)
    ps=d.Sections(1).PageSetup
    print('left margin=%.2fpt'%(ps.LeftMargin))
    for ti in (41,4,6,11):
        p=d.Paragraphs(ti); rng=p.Range
        fmt=p.Format
        print('--- i=%d: LeftIndent=%.2f FirstLineIndent=%.2f ---'%(ti,fmt.LeftIndent,fmt.FirstLineIndent))
        # first char x of each line: walk chars, detect y change
        prevy=None; lineno=0
        for k in range(1,min(rng.Characters.Count,200)+1):
            c=rng.Characters(k); y=c.Information(6); x=c.Information(5)
            if prevy is None or abs(y-prevy)>1.0:
                lineno+=1
                print('   line%d first char %r x=%.2f'%(lineno,c.Text.replace(chr(13),'CR'),x))
                if lineno>=3: break
            prevy=y
    d.Close(False)
finally:
    w.Quit()
