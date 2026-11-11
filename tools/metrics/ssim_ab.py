# -*- coding: utf-8 -*-
"""Correct, reusable SSIM A/B sentinel.

Renders every word_png base with DWrite under env A (var=1) and env B (default),
byte-compares, and SSIMs each changed doc's pages vs the cached Word reference
BOTH ways, reporting net (B - A). Uses the SAME skimage SSIM + RGB-load +
resize-to-match as the production pipeline (pipeline.ssim_calculator), so the
numbers match the gate. Earlier ad-hoc A/B scripts called calculate_ssim() with
the wrong signature (it takes dicts) and silently produced net=+0.0000 for
everything; this uses the low-level ssim directly.

Usage: python tools/metrics/ssim_ab.py OXI_S609_DISABLE   [base...] [--jobs=N]
       (A = OXI_S609_DISABLE=1, B = default. Optional base-prefix filter.)

Fast path (OXI_GDI_EXE set to a GDI renderer built from the SAME source as the
DWrite one): every document is first laid out under A and B by the GDI batch
mode with OXI_LAYOUT_FINGERPRINT (a hash of each page's full layout, every
element and field), and only documents whose fingerprints differ are drawn with
DWrite. Valid for changes in oxidocs-core (the layout); a change inside the
DWrite renderer's own drawing code is invisible to it, so gate that without
OXI_GDI_EXE. DWrite renders run --jobs at a time (default 4).
"""
import os, re, subprocess, sys, tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
# Resolve the repo root from this file's own location, not a hardcoded literal,
# so a git worktree gates against ITS OWN build/pipeline_data instead of silently
# borrowing the main tree's (the 2026-07-14/15 phantom-+5 footgun). From the main
# tree parents[2] == the old literal c:\Users\ryuji\oxi-main, so gate numbers are
# byte-identical there; from a worktree it resolves to the worktree root.
_REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(_REPO))
from pipeline.config import WORD_PNG_DIR, RENDER_DPI
from pipeline.ssim_calculator import _load_rgb, _resize_to_match
from skimage.metrics import structural_similarity as ssim
sys.stdout.reconfigure(encoding="utf-8")
REPO=str(_REPO)
# OXI_DWRITE_EXE lets a gate run against a SNAPSHOT of the renderer so a
# concurrent `cargo build --release` cannot swap the binary mid-run.
DW=os.environ.get("OXI_DWRITE_EXE") or os.path.join(REPO,"tools","oxi-dwrite-renderer","target","release","oxi-dwrite-renderer.exe")
DOCS=os.path.join(REPO,"tools","golden-test","documents","docx")
ARG=sys.argv[1] if len(sys.argv)>1 else "OXI_S609_DISABLE"
# "VAR" -> A sets VAR=1, B unsets. "VAR=VAL" -> A sets VAR=VAL, B unsets.
# COMMA-SEPARATED -> A sets them all, B unsets them all; that is how a bundle
# whose members only work together has to be gated (single-var callers are
# unaffected: a lone name has no comma).
AENV=[]
for part in ARG.split(","):
    if not part: continue
    if "=" in part: k,v=part.split("=",1)
    else: k,v=part,"1"
    AENV.append((k,v))
ENVVAR,AVAL=AENV[0]
JOBS=4
filt=[]
for a in sys.argv[2:]:
    if a.startswith("--jobs="): JOBS=max(1,int(a.split("=",1)[1]))
    else: filt.append(a)
GDI=os.environ.get("OXI_GDI_EXE")
def ssim2(wpng,opng):
    w=_load_rgb(wpng); o=_resize_to_match(_load_rgb(opng),w)
    return ssim(w,o,full=False,channel_axis=2,data_range=255)
bases=sorted({re.sub(r"_p\d+$","",n) for n in os.listdir(WORD_PNG_DIR)})
if filt: bases=[b for b in bases if any(b.startswith(f) for f in filt)]
def find(base):
    e=Path(DOCS)/(base+".docx")
    if e.exists(): return os.path.abspath(str(e))
    c=sorted(p for p in Path(DOCS).glob(base.split("_")[0]+"*.docx") if not p.name.startswith("~$"))
    return os.path.abspath(str(c[0])) if c else None
def render(docx,a,outdir):
    env=dict(os.environ)
    for k,v in AENV:
        if a: env[k]=v
        else: env.pop(k,None)
    Path(outdir).mkdir(parents=True,exist_ok=True)
    subprocess.run([DW,docx,str(Path(outdir)/"p"),str(RENDER_DPI)],capture_output=True,timeout=300,env=env)
    ps=[];i=1
    while (Path(outdir)/f"p_p{i}.png").exists(): ps.append(str(Path(outdir)/f"p_p{i}.png"));i+=1
    return ps
def fingerprints(items,a,tmp):
    """{docx: fingerprint text or None} from GDI --batch runs, JOBS processes."""
    env=dict(os.environ); env["OXI_LAYOUT_FINGERPRINT"]="1"
    for k,v in AENV:
        if a: env[k]=v
        else: env.pop(k,None)
    tag="A" if a else "B"
    chunks=[items[i::JOBS] for i in range(JOBS)]
    def run(ci):
        out={}
        lst=Path(tmp)/f"fp_{tag}_{ci}.txt"
        dumps={d:str(Path(tmp)/f"fp_{tag}_{ci}_{n}.json") for n,(_,d) in enumerate(chunks[ci])}
        lst.write_text("".join(f"{d}\t{dumps[d]}\n" for _,d in chunks[ci]),encoding="utf-8")
        subprocess.run([GDI,"--batch="+str(lst)],capture_output=True,env=env)
        for _,d in chunks[ci]:
            try: out[d]=Path(dumps[d]).read_text(encoding="utf-8")
            except OSError: out[d]=None
        return out
    res={}
    with ThreadPoolExecutor(JOBS) as ex:
        for r in ex.map(run,range(len(chunks))): res.update(r)
    return res
changed=[];checked=0
with tempfile.TemporaryDirectory() as tmp:
    seen=set(); items=[]
    for base in bases:
        d=find(base)
        if not d or d in seen: continue
        seen.add(d); items.append((base,d))
    checked=len(items)
    if GDI:
        fa=fingerprints(items,True,tmp); fb=fingerprints(items,False,tmp)
        # a GDI binary older than OXI_LAYOUT_FINGERPRINT writes its ordinary dump,
        # which leaves out colours and borders: refuse rather than under-report
        stale=[t for t in list(fa.values())+list(fb.values()) if t and not re.fullmatch(r'\{"pages":\[("[0-9a-f]{16}",?)*\]\}',t)]
        if stale: sys.exit("OXI_GDI_EXE does not write layout fingerprints (built before OXI_LAYOUT_FINGERPRINT); rebuild it or unset OXI_GDI_EXE")
        # a document that failed to lay out is drawn: no fingerprint, no skip
        todo=[(b,d) for b,d in items if fa.get(d) is None or fa.get(d)!=fb.get(d)]
        print(f"layout fingerprints: {len(todo)} of {checked} differ (or failed) -> drawn with DWrite")
    else:
        todo=items
    def ab(item):
        base,d=item
        ad=Path(tmp)/"A"/Path(d).stem; bd=Path(tmp)/"B"/Path(d).stem
        pa=render(d,True,ad); pb=render(d,False,bd)
        diff=(len(pa)!=len(pb)) or any(open(x,"rb").read()!=open(y,"rb").read() for x,y in zip(pa,pb))
        return (base,ad,bd,len(pa),len(pb)) if diff else None
    with ThreadPoolExecutor(JOBS) as ex:
        changed=[r for r in ex.map(ab,todo) if r]
    print(f"A={','.join(k+'='+v for k,v in AENV)} vs B=default | checked {checked}; {len(changed)} changed bytes")
    tot=0.0; wins=losses=0; details=[]
    for base,ad,bd,na,nb in changed:
        wdir=Path(WORD_PNG_DIR)/base; i=1; net=0.0; npg=0
        while True:
            wp=wdir/f"page_{i:04d}.png"
            if not wp.exists(): break
            ap=Path(ad)/f"p_p{i}.png"; bp=Path(bd)/f"p_p{i}.png"
            if not ap.exists() or not bp.exists(): break
            try: sa=ssim2(str(wp),str(ap)); sb=ssim2(str(wp),str(bp))
            except Exception: i+=1; continue
            net+=(sb-sa); npg+=1; i+=1
        if npg==0: 
            details.append((base,None,na,nb)); continue
        tot+=net
        if net>0.0005: wins+=1
        elif net<-0.0005: losses+=1
        details.append((base,net,na,nb))
    for base,net,na,nb in sorted(details,key=lambda x:(x[1] if x[1] is not None else 0)):
        if net is None: print(f"  {base}: NO word_png pages found (pages {na}/{nb})")
        else:
            flag=" <<< REGRESS" if net<-0.0005 else (" >>> improve" if net>0.0005 else "")
            print(f"  {base}: pages={na}/{nb} net(B-A)={net:+.4f}{flag}")
    print(f"\nTOTAL net(B-A)={tot:+.4f}; improved {wins}, regressed {losses}")
