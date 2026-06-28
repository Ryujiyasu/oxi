# -*- coding: utf-8 -*-
import os, sys, json, subprocess
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
RENDERER=os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
DOCX=os.path.abspath('tools/golden-test/documents/docx/tokyoshugyo_000599795.docx')
def norm(s): return s.replace("　","").replace(" ","").strip()
W=json.load(open("C:/tmp/tks_word_glyphs.json",encoding="utf-8"))
wfirst={}
for pi,pg in enumerate(W["pages"]):
    ys=defaultdict(list)
    for g in pg["glyphs"]:
        if g["y"]>=88 and g["y"]<785: ys[round(g["y"],0)].append(g)
    if ys:
        y0=min(ys); gg=sorted(ys[y0],key=lambda g:g["x"]); wfirst[pi+1]=norm("".join(g["char"] for g in gg))
def render_count(env):
    e=dict(os.environ); e.update(env)
    dump='C:/tmp/tks_joint.json'
    subprocess.run([RENDERER,DOCX,'C:/tmp/tks_j','96','--dump-layout='+dump],env=e,capture_output=True)
    od=json.load(open(dump,encoding="utf-8")); of={}
    for pg in od["pages"]:
        rows=defaultdict(lambda:['',0])
        for el in pg["elements"]:
            if el.get("type")=="text" and el.get("text","").strip() and round(el["y"],1)>=88:
                y=round(el["y"],1); rows[y][0]+=el["text"]
        if rows:
            y0=min(rows); of[pg["page"]]=norm(rows[y0][0])
    npages=len(od["pages"]); ndiv=0
    for p in range(1,max(npages,len(W["pages"]))+1):
        o=of.get(p,""); w=wfirst.get(p,"")
        if o[:10]!=w[:10]: ndiv+=1
    return npages, ndiv
for label,env in [("default",{}),("S590(body)",{"OXI_S590":"1"}),
                  ("S590+S586(cell44)",{"OXI_S590":"1","OXI_S586":"1"}),
                  ("S589(body-natural)",{"OXI_S589":"1"})]:
    np,nd=render_count(env)
    print(f"{label:<22} oxi_pages={np:>3}  page-top divergences={nd}")

print("\n=== cap sweep under S590 (OXI_S475_SOLO, pair=6.0) ===")
for solo in ["1.0","1.5","2.0","2.5","3.0"]:
    np,nd=render_count({"OXI_S590":"1","OXI_S475_SOLO":solo})
    print(f"  solo={solo:<5} oxi_pages={np:>3}  page-top divergences={nd}")
