import win32com.client as win32, statistics
HPOS=5;VPOS=6
D=r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\chargrid_wrap\cg_cell_12pt.docx"
word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
doc=word.Documents.Open(D,ReadOnly=True)
try:
    t=doc.Tables(1); c=t.Cell(1,1).Range
    n=c.Characters.Count
    xs=[];ys=[]
    for i in range(min(n,60)):
        ch=doc.Range(c.Start+i,c.Start+i+1)
        xs.append(round(ch.Information(HPOS),2)); ys.append(round(ch.Information(VPOS),1))
    y0=ys[0]; first_line=sum(1 for y in ys if y==y0)
    adv=[round(xs[i+1]-xs[i],2) for i in range(min(8,len(xs)-1)) if ys[i]==ys[i+1]==y0]
    print(f"Word CELL 12pt: chars on line1 = {first_line}  advances {adv} median {statistics.median(adv) if adv else '?'}")
    print(f"  cell width 9000tw=450pt; S466 wrap adv=12.405, emit adv=12.355; OFF emit=12.575")
finally:
    doc.Close(False);word.Quit()
