# -*- coding: utf-8 -*-
import json, fitz, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
od=json.load(open("C:/tmp/tks.json",encoding="utf-8"))
def norm(s): return s.replace("　","").replace(" ","").strip()
ofirst={}; opages=len(od["pages"])
for pg in od["pages"]:
    rows=defaultdict(list)
    for e in pg["elements"]:
        if e.get("type")=="text" and e.get("text","").strip() and round(e["y"],1)>=88:
            rows[round(e["y"],1)].append(e)
    if rows:
        y0=min(rows); ofirst[pg["page"]]=norm("".join(c["text"] for c in sorted(rows[y0],key=lambda c:c["x"])))
doc=fitz.open(r"C:\tmp\tokyoshugyo_word.pdf")
wfirst={}; wpages=doc.page_count
# detect header text (appears on most pages at top) — skip the topmost if it repeats
for pi in range(wpages):
    L=[]
    for blk in doc.load_page(pi).get_text("rawdict")["blocks"]:
        if blk.get("type")!=0: continue
        for ln in blk.get("lines",[]):
            cs=[c for sp in ln.get("spans",[]) for c in sp.get("chars",[])]
            t="".join(c["c"] for c in cs).strip()
            if t and round(ln["bbox"][1],1)>=88: L.append((round(ln["bbox"][1],1),norm(t)))
    L.sort()
    if L: wfirst[pi+1]=L[0][1]
print(f"Oxi pages={opages}  Word pages={wpages}")
print("page-top first-body-line: divergences (Oxi vs Word):")
ndiv=0
for p in range(1,max(opages,wpages)+1):
    o=ofirst.get(p,""); w=wfirst.get(p,"")
    if o[:12]!=w[:12]:
        ndiv+=1
        if ndiv<=20: print(f"  p{p:3} O:'{o[:26]}'  W:'{w[:26]}'")
print(f"\nTOTAL page-top divergences: {ndiv}/{max(opages,wpages)}")
