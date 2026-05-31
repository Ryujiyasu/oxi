import win32com.client as win32, statistics
HPOS=5;VPOS=6
D=r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\chargrid_wrap\cg_mincho_12_in_105grid.docx"
word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
doc=word.Documents.Open(D,ReadOnly=True)
try:
    rng=doc.Paragraphs(1).Range
    xs=[];ys=[]
    for i in range(8):
        c=doc.Range(rng.Start+i,rng.Start+i+1); xs.append(round(c.Information(HPOS),3)); ys.append(round(c.Information(VPOS),1))
    adv=[round(xs[i+1]-xs[i],3) for i in range(len(xs)-1) if ys[i]==ys[i+1]==ys[0]]
    print("12pt char (in 10.5 default grid) Word advances:",adv,"median",statistics.median(adv) if adv else "?")
    print("  natural 12pt MS Mincho fullwidth = 12.0 ; S466 expected_w = 12*pitch/10.5 = 12*1.0338 =",round(12*10.854736/10.5,3))
finally:
    doc.Close(False);word.Quit()
