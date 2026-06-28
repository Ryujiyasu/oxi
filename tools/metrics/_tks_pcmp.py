import json, fitz, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
P=int(sys.argv[1])
def norm(s): return s.replace("　","").replace(" ","").strip()
od=json.load(open("C:/tmp/tks.json",encoding="utf-8"))
ol=[]
for pg in od["pages"]:
    if pg["page"]!=P: continue
    rows=defaultdict(list)
    for e in pg["elements"]:
        if e.get("type")=="text" and e.get("text","").strip() and round(e["y"],1)>=88: rows[round(e["y"],1)].append(e)
    for y in sorted(rows): ol.append(norm("".join(c["text"] for c in sorted(rows[y],key=lambda c:c["x"]))))
doc=fitz.open(r"C:\tmp\tokyoshugyo_word.pdf")
wl=[]
for blk in doc.load_page(P-1).get_text("rawdict")["blocks"]:
    if blk.get("type")!=0: continue
    for ln in blk.get("lines",[]):
        cs=[c for sp in ln.get("spans",[]) for c in sp.get("chars",[])]
        t=norm("".join(c["c"] for c in cs))
        if t and round(ln["bbox"][1],1)>=88: wl.append((round(ln["bbox"][1],1),t))
wl=[t for _,t in sorted(wl)]
print(f"p{P}: Oxi {len(ol)} lines | Word {len(wl)} lines")
for i in range(max(len(ol),len(wl))):
    o=ol[i] if i<len(ol) else ""; w=wl[i] if i<len(wl) else ""
    mark="" if o[:10]==w[:10] else " <<<"
    print(f"  {o[:32]:<34}| {w[:32]}{mark}")
