# -*- coding: utf-8 -*-
# Per-Word-page delta map: for each Word page, which Oxi page does its first-line
# content head appear on (page_delta = oxi - word). Reveals the cascade structure.
import os, sys, json, subprocess
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
RENDERER=os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
DOCX=os.path.abspath('tools/golden-test/documents/docx/tokyoshugyo_000599795.docx')
def norm(s): return s.replace("　","").replace(" ","").strip()
W=json.load(open("C:/tmp/tks_word_glyphs.json",encoding="utf-8"))
# Word: collect ALL line heads per page (sorted by y), keyed page->list of normed line texts
wlines={}
for pi,pg in enumerate(W["pages"]):
    ys=defaultdict(list)
    for g in pg["glyphs"]:
        if g["y"]>=88 and g["y"]<785: ys[round(g["y"],0)].append(g)
    lines=[]
    for y0 in sorted(ys):
        gg=sorted(ys[y0],key=lambda g:g["x"]); lines.append(norm("".join(g["char"] for g in gg)))
    wlines[pi+1]=lines

def render(env):
    e=dict(os.environ); e.update(env)
    dump='C:/tmp/tks_dm.json'
    subprocess.run([RENDERER,DOCX,'C:/tmp/tks_dm','96','--dump-layout='+dump],env=e,capture_output=True)
    od=json.load(open(dump,encoding="utf-8"))
    olines={}  # oxi page -> list of normed line texts (sorted by y)
    for pg in od["pages"]:
        rows=defaultdict(str)
        for el in pg["elements"]:
            if el.get("type")=="text" and el.get("text","").strip() and round(el["y"],1)>=88:
                rows[round(el["y"],1)]+=el["text"]
        olines[pg["page"]]=[norm(rows[y]) for y in sorted(rows)]
    return len(od["pages"]), olines

def deltamap(env, label, show=False):
    npages, olines = render(env)
    # build a flat list of (oxi_page, linehead) for searching
    flat=[]
    for op in sorted(olines):
        for ln in olines[op]:
            if ln: flat.append((op,ln))
    deltas=defaultdict(int); rows=[]
    cursor=0  # monotonic index into flat (pages are in order; avoids TOC re-match)
    for wp in range(1,len(W["pages"])+1):
        wh = wlines.get(wp,[""])[0] if wlines.get(wp) else ""
        if not wh or len(wh)<4:
            rows.append((wp,wh,"?",None)); continue
        # monotonic forward search from `cursor` for the matching oxi line
        found=None; fidx=None
        for i in range(cursor,len(flat)):
            op,ln=flat[i]
            if ln[:8]==wh[:8]:
                found=op; fidx=i; break
        if found is None:
            rows.append((wp,wh,"MISS",None)); deltas["MISS"]+=1
        else:
            cursor=fidx+1
            d=found-wp; deltas[d]+=1; rows.append((wp,wh,found,d))
    summ=", ".join(f"{k}:{deltas[k]}" for k in sorted(deltas,key=lambda x:(isinstance(x,str),x)))
    print(f"{label:<28} oxi_pages={npages:>3}  {{{summ}}}")
    if show:
        for wp,wh,found,d in rows:
            if d not in (0,None):
                print(f"   Wp{wp:>2} d={d:+d} oxi={found}  {wh[:22]}")
    return rows

if __name__=="__main__":
    show = "--show" in sys.argv
    combos=[("default",{}),
            ("S693",{"OXI_S693":"1"}),
            ]
    # allow extra env from CLI: KEY=VAL
    extra={}
    for a in sys.argv[1:]:
        if "=" in a and not a.startswith("--"):
            k,v=a.split("=",1); extra[k]=v
    if extra:
        combos.append((("+".join(f"{k}={v}" for k,v in extra.items()))[:28], extra))
    for label,env in combos:
        deltamap(env,label,show=show)
