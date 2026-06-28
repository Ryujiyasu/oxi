import json, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
def norm(s): return s.replace("　","").replace(" ","").strip()
W=json.load(open("C:/tmp/tks_word_glyphs.json",encoding="utf-8"))
wfirst={}
for pi,pg in enumerate(W["pages"]):
    ys=defaultdict(list)
    for g in pg["glyphs"]:
        if g["y"]>=88: ys[round(g["y"],0)].append(g)
    if ys:
        y0=min(ys); gg=sorted(ys[y0],key=lambda g:g["x"])
        wfirst[pi+1]=norm("".join(g["char"] for g in gg))
def odump(fn):
    od=json.load(open(fn,encoding="utf-8")); of={}
    for pg in od["pages"]:
        rows=defaultdict(list)
        for e in pg["elements"]:
            if e.get("type")=="text" and e.get("text","").strip() and round(e["y"],1)>=88:
                rows[round(e["y"],1)].append(e)
        if rows:
            y0=min(rows); of[pg["page"]]=norm("".join(c["text"] for c in sorted(rows[y0],key=lambda c:c["x"])))
    return of, len(od["pages"])
for lbl,fn in [("K=0","C:/tmp/tks_dump0.json"),("K=0.85","C:/tmp/tks_dump85.json")]:
    of,op=odump(fn); ndiv=0; divs=[]
    for p in range(1,max(op,len(W["pages"]))+1):
        o=of.get(p,""); w=wfirst.get(p,"")
        if o[:10]!=w[:10]: ndiv+=1; divs.append((p,o[:20],w[:20]))
    print(f"[{lbl}] Oxi pages={op}  divergences={ndiv}")
    for p,o,w in divs: print(f"  p{p:3} O:'{o}'  W:'{w}'")
    print()
