"""Measure Word's EXACT charGrid char advance via COM horizontal positions.
Determines if Oxi's 11.075pt advance over-expands (pitch error) or is a pure
wrap-boundary epsilon. wdHorizontalPositionRelativeToPage = 5."""
import win32com.client as win32
HPOS=5;VPOS=6
REPRO=r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\chargrid_wrap"
word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
for nm,fs in [("cg_mincho_10p5",10.5),("cg_mincho_10",10.0),("cg_mincho_9",9.0)]:
    doc=word.Documents.Open(f"{REPRO}\{nm}.docx",ReadOnly=True)
    try:
        rng=doc.Paragraphs(1).Range
        # x of char 1, char 2, ... up to ~6 on line 1 (same Y)
        xs=[];ys=[]
        for i in range(0,8):
            c=doc.Range(rng.Start+i,rng.Start+i+1)
            xs.append(round(c.Information(HPOS),3)); ys.append(round(c.Information(VPOS),1))
        # advances on line 1
        adv=[round(xs[i+1]-xs[i],3) for i in range(len(xs)-1) if ys[i]==ys[i+1]==ys[0]]
        import statistics
        print(f"{nm} fs{fs}: Word char x positions {xs[:6]}")
        print(f"   Word advances: {adv}  median={statistics.median(adv) if adv else '?'}  (Oxi S466=11.075/10.368/9.371)")
    finally:
        doc.Close(False)
word.Quit()
