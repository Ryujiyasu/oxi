# -*- coding: utf-8 -*-
# Per-LINE monotonic page-delta over the whole doc; prints every transition point
# (where Oxi's page-delta vs Word changes) = the cascade origins.
import os, sys, json, subprocess
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
RENDERER=os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
DOCX=os.path.abspath('tools/golden-test/documents/docx/tokyoshugyo_000599795.docx')
def norm(s): return s.replace("　","").replace(" ","").strip()
W=json.load(open("C:/tmp/tks_word_glyphs.json",encoding="utf-8"))
# Word: flat list of (word_page, line_text) in reading order
wflat=[]
for pi,pg in enumerate(W["pages"]):
    ys=defaultdict(list)
    for g in pg["glyphs"]:
        if g["y"]>=88 and g["y"]<785: ys[round(g["y"],0)].append(g)
    for y0 in sorted(ys):
        gg=sorted(ys[y0],key=lambda g:g["x"]); t=norm("".join(g["char"] for g in gg))
        if len(t)>=5: wflat.append((pi+1,t))

def render(env):
    e=dict(os.environ); e.update(env)
    dump='C:/tmp/tks_drift.json'
    subprocess.run([RENDERER,DOCX,'C:/tmp/tks_dr','96','--dump-layout='+dump],env=e,capture_output=True)
    od=json.load(open(dump,encoding="utf-8"))
    oflat=[]
    for pg in od["pages"]:
        rows=defaultdict(str)
        for el in pg["elements"]:
            if el.get("type")=="text" and el.get("text","").strip() and round(el["y"],1)>=88:
                rows[round(el["y"],1)]+=el["text"]
        for y in sorted(rows):
            t=norm(rows[y])
            if len(t)>=5: oflat.append((pg["page"],t))
    return oflat, len(od["pages"])

def run(env,label, lo=0, hi=999):
    oflat,npages=render(env)
    cur=0; prev_d=None; trans=[]
    for wp,wt in wflat:
        found=None
        for i in range(cur,len(oflat)):
            op,ot=oflat[i]
            n=min(len(wt),len(ot))
            if n>=5 and wt[:n]==ot[:n]:
                found=(op,i); break
        if not found: continue
        op,i=found; cur=i+1; d=op-wp
        if d!=prev_d:
            trans.append((wp,op,d,wt[:26])); prev_d=d
    print(f"=== {label}  oxi_pages={npages} ===")
    for wp,op,d,t in trans:
        if lo<=wp<=hi:
            print(f"  Wp{wp:>2}->Oxi{op:>2} d={d:+d}  {t}")

if __name__=="__main__":
    lo,hi=0,999
    env={}
    for a in sys.argv[1:]:
        if a.startswith("--range="):
            lo,hi=map(int,a.split("=",1)[1].split(":"))
        elif "=" in a: k,v=a.split("=",1); env[k]=v
    run(env, "+".join(f"{k}={v}" for k,v in env.items()) or "default", lo, hi)
