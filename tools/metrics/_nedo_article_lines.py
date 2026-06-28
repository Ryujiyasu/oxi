# -*- coding: utf-8 -*-
import json, fitz, re, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
PDF=r"C:\tmp\nedocontract_word.pdf"; HEADER="一般再委託用"
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
            if t and not t.startswith(HEADER): out.append(t)
    return out
def oxi_lines():
    d=json.load(open("C:/tmp/n_def.json",encoding="utf-8")); out=[]
    for pg in d["pages"]:
        rows=defaultdict(list)
        for e in pg["elements"]:
            if e.get("type")=="text": rows[round(e["y"],1)].append(e)
        for y in sorted(rows):
            t="".join(c["text"] for c in sorted(rows[y],key=lambda c:c["x"])).strip()
            if t and not t.startswith(HEADER): out.append(t)
    return out
wl=word_lines(); ol=oxi_lines()
# anchors = lines starting with 第N条 (article headings) - unique
art=re.compile(r'^第[０-９0-9一二三四五六七八九十]+条(　|s|$)')
def anchors(lines):
    return [(i,l) for i,l in enumerate(lines) if art.match(l) and 'から' not in l[:12] and 'まで' not in l[:12]]
wa=anchors(wl); oa=anchors(ol)
print(f"Word lines={len(wl)} Oxi lines={len(ol)}; Word articles={len(wa)} Oxi articles={len(oa)}")
# match articles by their heading text prefix
wmap={l[:8]:i for i,l in wa}
omap={l[:8]:i for i,l in oa}
common=[k for k in omap if k in wmap]
# for consecutive common anchors, count lines between
def seg_counts(anchor_idx_map, lines, keys):
    pass
# build per-article line spans
def spans(amap, total):
    items=sorted(amap.items(), key=lambda kv: kv[1])
    out={}
    for j,(k,idx) in enumerate(items):
        nxt = items[j+1][1] if j+1<len(items) else total
        out[k]=(idx,nxt)
    return out
ws=spans(wmap,len(wl)); os_=spans(omap,len(ol))
print("\nArticles where Oxi line-count != Word (the under/over-fit location):")
for k in sorted(common, key=lambda k: ws[k][0]):
    wn=ws[k][1]-ws[k][0]; on=os_[k][1]-os_[k][0]
    if wn!=on:
        head=[l for i,l in wa if l[:8]==k][0][:14]
        print(f"  {head:16} Word {wn:3} lines / Oxi {on:3}  delta {on-wn:+d}")
