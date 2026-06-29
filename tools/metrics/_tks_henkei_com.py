# -*- coding: utf-8 -*-
# COM-measure the real tokyoshugyo 変形 region (注 → empties → 〔例３〕) per-paragraph
# Y gaps (Information(6) with collapsed start range, R30 fix). Compare to Oxi.
import os, sys
sys.stdout.reconfigure(encoding="utf-8")
DOCX=os.path.abspath('tools/golden-test/documents/docx/tokyoshugyo_000599795.docx')
import win32com.client as w
app=w.Dispatch('Word.Application'); app.Visible=False
doc=app.Documents.Open(DOCX, ReadOnly=True)
n=doc.Paragraphs.Count
# find the body 〔例３〕 paragraph (the heading, not TOC). TOC 〔例... won't match this exact.
target=None
for i in range(1,n+1):
    t=doc.Paragraphs(i).Range.Text
    if t.startswith('〔例３〕') and '規程例' in t:
        target=i; break
if not target:
    print("〔例３〕 not found"); doc.Close(False); app.Quit(); sys.exit()
print(f"〔例３〕 at COM para {target} of {n}")
prev=None
for i in range(target-14, target+3):
    rng=doc.Paragraphs(i).Range
    y=doc.Range(rng.Start,rng.Start).Information(6)
    pg=doc.Range(rng.Start,rng.Start).Information(3)  # page number
    txt=rng.Text.replace('\r','').replace('\x07','')[:20]
    gap=(y-prev) if prev is not None else 0.0
    mark='E' if not txt.strip() else ' '
    print(f"  para{i:>4} pg={pg} y={y:7.2f} gap={gap:6.2f} {mark} {txt!r}")
    prev=y
doc.Close(False); app.Quit()
