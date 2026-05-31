"""Measure Word's actual char advance for de6e32 p7 list items (table cell, 12pt)
to decide if Word expands cell chars to grid (=S466 right) or uses natural (=OFF right)."""
import win32com.client as win32, statistics
HPOS=5;VPOS=6;PAGE=3
D=r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\de6e32b5960b_tokumei_08_01-1.docx"
word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
doc=word.Documents.Open(D,ReadOnly=True)
try:
    # find a paragraph on page 7 with CJK content (list item)
    found=0
    for p in doc.Paragraphs:
        rng=p.Range; st=doc.Range(rng.Start,rng.Start)
        if st.Information(PAGE)!=7: continue
        t=p.Range.Text.strip()
        if len(t)<6: continue
        # measure advances of first 8 chars
        xs=[];ys=[];fs=p.Range.Font.Size
        for i in range(8):
            c=doc.Range(rng.Start+i,rng.Start+i+1)
            xs.append(round(c.Information(HPOS),2)); ys.append(round(c.Information(VPOS),1))
        adv=[round(xs[i+1]-xs[i],2) for i in range(len(xs)-1) if ys[i]==ys[i+1]==ys[0]]
        if adv:
            print(f"p7 para fs={fs} '{t[:12]}': Word advances {adv} median {statistics.median(adv)}")
            found+=1
        if found>=4: break
    print(f"\n  natural fs=12 -> 12.0 ; S466 (fs=12, default10.5) -> 12.405 ; body-12pt Word measured 12.375")
finally:
    doc.Close(False);word.Quit()
