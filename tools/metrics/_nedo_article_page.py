# -*- coding: utf-8 -*-
import json, fitz, re, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
PDF=r"C:\tmp\nedocontract_word.pdf"; HEADER="一般再委託用"
ART=re.compile(r'^第[０-９0-9一二三四五六七八九十]+条')
def word_lines():
    doc=fitz.open(PDF); out=[]
    for pi in range(doc.page_count):
        L=[]
        for blk in doc.load_page(pi).get_text("rawdict")["blocks"]:
            if blk.get("type")!=0: continue
            for ln in blk.get("lines",[]):
                cs=[(c["c"],c["bbox"][0]) for sp in ln.get("spans",[]) for c in sp.get("chars",[])]
                if cs: cs.sort(key=lambda t:t[1]); L.append((round(ln["bbox"][1],1),"".join(c[0] for c in cs).strip()))
        L.sort()
        for _,t in L:
            if t and not t.startswith(HEADER): out.append((pi+1,t))
    return out
def oxi_lines():
    d=json.load(open("C:/tmp/n_def.json",encoding="utf-8")); out=[]
    for pg in d["pages"]:
        rows=defaultdict(list)
        for e in pg["elements"]:
            if e.get("type")=="text": rows[round(e["y"],1)].append(e)
        for y in sorted(rows):
            t="".join(c["text"] for c in sorted(rows[y],key=lambda c:c["x"])).strip()
            if t and not t.startswith(HEADER): out.append((pg["page"],t))
    return out
wl=word_lines(); ol=oxi_lines()
# article heading = 第N条 NOT followed by から/まで (references)
def heads(lines):
    h={}
    for pg,t in lines:
        if ART.match(t) and 'から' not in t[:10] and 'まで' not in t[:10]:
            key=ART.match(t).group()
            if key not in h: h[key]=pg
    return h
wh=heads(wl); oh=heads(ol)
common=[k for k in oh if k in wh]
print(f"Word headings={len(wh)} Oxi headings={len(oh)} common={len(common)}")
print("Articles where Oxi START PAGE != Word START PAGE (the REAL pagination diff):")
diffs=0
for k in sorted(common, key=lambda k: wh[k]):
    if oh[k]!=wh[k]:
        diffs+=1; print(f"  {k:14} Word p{wh[k]:2} / Oxi p{oh[k]:2}  delta {oh[k]-wh[k]:+d}")
print(f"\nTOTAL articles with page mismatch: {diffs}/{len(common)}")
