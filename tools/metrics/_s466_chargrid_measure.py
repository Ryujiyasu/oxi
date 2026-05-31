import json
import win32com.client as win32
VPOS=6
REPRO=r"C:\Users\ryuji\oxi-main\tools\golden-test\repros\chargrid_wrap"
NAMES=["cg_mincho_10p5","cg_mincho_10","cg_mincho_9"]

def oxi_first_line_chars(dump):
    d=json.load(open(dump,encoding="utf-8"))
    els=[e for pg in d["pages"] if pg["page"]==1 for e in pg["elements"] if e["type"]=="text" and (e.get("text") or "").strip()]
    if not els: return None,None
    y0=min(e["y"] for e in els)
    # chars on first line = sum len of text of elements at y0
    first=sum(len(e["text"]) for e in els if abs(e["y"]-y0)<1.0)
    # total lines = distinct y
    nlines=len(set(round(e["y"],1) for e in els))
    return first,nlines

word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
print(f"{'repro':<16}{'Wchars/L1':>10}{'Ochars/L1':>10}{'Wlines':>8}{'Olines':>8}")
for nm in NAMES:
    docx=f"{REPRO}\{nm}.docx"
    doc=word.Documents.Open(docx,ReadOnly=True)
    try:
        rng=doc.Paragraphs(1).Range
        n=rng.Characters.Count
        # walk chars, detect first Y change
        prevY=None; first_line=0; ys=[]
        for i in range(1,min(n,120)+1):
            ch=doc.Range(rng.Start+i-1, rng.Start+i)
            y=round(ch.Information(VPOS),1)
            ys.append(y)
            if prevY is None: prevY=y
            if y==prevY: first_line=i
            elif y!=prevY and first_line and len([1 for yy in ys if yy==prevY]):
                pass
        # count chars on the first (minimum y) line
        y0=min(ys); wfirst=sum(1 for yy in ys if yy==y0)
        wlines=len(set(ys))
    finally:
        doc.Close(False)
    ofirst,olines=oxi_first_line_chars(f"C:/tmp/cg_{nm}/layout.json")
    print(f"{nm:<16}{wfirst:>10}{str(ofirst):>10}{wlines:>8}{str(olines):>8}")
word.Quit()
