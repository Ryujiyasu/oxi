# -*- coding: utf-8 -*-
"""Measure 、/。 advance for several jc=LEFT numbered-list paras in harassmanual.
If always ~6pt -> inherent half-width yakumono; if 12pt except tight lines ->
oikomi-only (line-fit driven)."""
import os, win32com.client as win32
DOCX=os.path.abspath("tools/golden-test/documents/docx/harassmanual_001466344.docx")
w=win32.gencache.EnsureDispatch("Word.Application");w.Visible=False
d=w.Documents.Open(DOCX,ReadOnly=True)
try:
    cnt=0
    for i in range(1,d.Paragraphs.Count+1):
        r=d.Paragraphs(i).Range; t=r.Text.rstrip("\r\n")
        jc=r.ParagraphFormat.Alignment  # 0=left,1=center,3=justify
        if jc!=0: continue   # jc=left only
        # measure each 、。 in this para + its display line-number
        for k in range(len(t)-1):
            if t[k] in "、。":
                x1=d.Range(r.Start+k,r.Start+k+1).Information(5)
                x2=d.Range(r.Start+k+1,r.Start+k+2).Information(5)
                adv=x2-x1
                # is this near line end? check line number of this char vs next
                ln1=d.Range(r.Start+k,r.Start+k+1).Information(10)
                ln2=d.Range(r.Start+k+1,r.Start+k+2).Information(10)
                pos="LINE-END" if ln1!=ln2 else "mid"
                print(f"para{i} jc=L  {t[k]!r} adv={adv:.2f} ({pos})  ctx={t[max(0,k-3):k+2]!r}")
                cnt+=1
                if cnt>=18: break
        if cnt>=18: break
finally:
    d.Close(False);w.Quit()
