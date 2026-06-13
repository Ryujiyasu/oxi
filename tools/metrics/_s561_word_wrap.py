# -*- coding: utf-8 -*-
"""Word COM: per-char x/y of the (5)裁量 paragraph to find the exact wrap point
(which char first overflows Word's line-1 budget) + the line-1 end x (= Word's
effective cell content width). Find-based (vMerge blocks Cell access)."""
import sys,os
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8')
wd=win32.gencache.EnsureDispatch('Word.Application'); wd.Visible=False
path=os.path.abspath('tools/golden-test/documents/docx/roudoujoken_001161383.docx')
doc=wd.Documents.Open(path,ReadOnly=True)
# find the (5)裁量 paragraph
rng=doc.Content
f=rng.Find; f.ClearFormatting(); f.Text='裁量労働制'
if f.Execute():
    para=rng.Paragraphs(1).Range
    print('para text: %r'%para.Text.strip()[:60])
    print('per-char (idx, char, x=Info5, y=Info6):')
    prev_y=None; line=1
    s=para.Text
    for i in range(len(s)):
        ch=doc.Range(para.Start+i, para.Start+i+1)
        c=s[i]
        if c in '\r\n\x07': continue
        x=ch.Information(5); y=ch.Information(6)
        mark=''
        if prev_y is not None and abs(y-prev_y)>3:
            line+=1; mark='  <<< LINE %d (wrap before %r)'%(line,c)
        print('  %2d %r x=%6.1f y=%6.1f%s'%(i,c,x,y,mark))
        prev_y=y
doc.Close(False); wd.Quit()
