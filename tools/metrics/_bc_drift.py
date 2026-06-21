# bunkacontract: per-paragraph Y (Word COM Info6) vs page, to profile the drift.
import win32com.client as w, glob
DOC=glob.glob(r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\*bunka*.docx")[0]
app=w.Dispatch("Word.Application"); app.Visible=False
try:
    d=app.Documents.Open(DOC, ReadOnly=True)
    P=d.Paragraphs; n=P.Count
    def y(rng): return d.Range(rng.Start,rng.Start).Information(6)
    def pg(rng): return d.Range(rng.Start,rng.Start).Information(3)
    prevpg=None
    for i in range(1,min(n+1,60)):
        p=P(i); rng=p.Range
        txt=(rng.Text or "").replace("\r","").replace("\x07","")[:16]
        yy=y(rng); pgn=pg(rng)
        # footnote ref?
        hasfn = rng.Footnotes.Count>0
        mark = " <FN>" if hasfn else ""
        pb = " ===PAGE %d==="%pgn if pgn!=prevpg else ""
        print(f"  i={i:3d} p{pgn} y={yy:7.2f}{mark}{pb} :: {txt!r}")
        prevpg=pgn
    d.Close(False)
finally: app.Quit()
