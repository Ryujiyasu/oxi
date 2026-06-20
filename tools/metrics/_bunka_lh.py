import win32com.client as w, glob
DOC = glob.glob(r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\*bunka*.docx")[0]
app=w.Dispatch("Word.Application"); app.Visible=False
try:
    d=app.Documents.Open(DOC, ReadOnly=True)
    P=d.Paragraphs; n=P.Count
    def y(rng): return d.Range(rng.Start,rng.Start).Information(6)
    def pg(rng): return d.Range(rng.Start,rng.Start).Information(3)
    # body font
    fc=d.Range(P(5).Range.Start, P(5).Range.Start+1)
    print("compat:", d.CompatibilityMode)
    print("body font FE:", fc.Font.NameFarEast, "asc:", fc.Font.Name, "sz:", fc.Font.Size)
    # find a long multi-line body paragraph; measure intra-line gap via two consecutive
    # paragraphs that are single-line, same style, no spacing
    prev=None; rows=[]
    for i in range(1,min(n+1,60)):
        p=P(i); rng=p.Range
        txt=(rng.Text or "").replace("\r","").replace("\x07","")
        rows.append((i,pg(rng),round(y(rng),2),p.Format.SpaceBefore,p.Format.SpaceAfter,p.Format.LineSpacingRule,round(p.Format.LineSpacing,2),txt[:20]))
    for i,pgn,yy,sb,sa,lsr,ls,txt in rows:
        print(f"  i={i:3d} p{pgn} y={yy:7.2f} sb={sb} sa={sa} lsr={lsr} ls={ls} :: {txt!r}")
    d.Close(False)
finally:
    app.Quit()
