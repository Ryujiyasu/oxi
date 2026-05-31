"""Definitive S463 effect: compare ON vs OFF Oxi-SSIM scores directly,
bypassing the stored baseline. ON = /c/tmp/s463_ON_scores.json (rendered with
OXI_S463_LATIN_BORDER=1), OFF = pipeline_data/ssim_current_post_s160.json
(rendered without the env var). Both scored vs the same Word PNGs."""
import json
ON=json.load(open(r"C:/tmp/s463_ON_scores.json",encoding="utf-8"))
OFF=json.load(open("pipeline_data/ssim_current_post_s160.json",encoding="utf-8"))
rows=[]
for doc,pages in OFF.items():
    for pg,off in pages.items():
        on=ON.get(doc,{}).get(pg)
        if on is None: continue
        rows.append((doc,pg,off,on,on-off))
on_all=[r[3] for r in rows]; off_all=[r[2] for r in rows]
mean=lambda x:sum(x)/len(x)
print(f"pages compared: {len(rows)}")
print(f"OFF mean: {mean(off_all):.4f}   ON mean: {mean(on_all):.4f}   delta: {mean(on_all)-mean(off_all):+.4f}")
def buckets(x):
    return f">=0.99:{sum(1 for v in x if v>=0.99)} >=0.95:{sum(1 for v in x if v>=0.95)} >=0.90:{sum(1 for v in x if v>=0.90)} <0.70:{sum(1 for v in x if v<0.70)}"
print("OFF buckets:",buckets(off_all))
print("ON  buckets:",buckets(on_all))
imp=[r for r in rows if r[4]>0.002]; reg=[r for r in rows if r[4]<-0.002]
print(f"\nimproved(>+0.002): {len(imp)}   regressed(<-0.002): {len(reg)}   unchanged: {len(rows)-len(imp)-len(reg)}")
# bottom-N floor (Phase-3 gate)
off_sorted=sorted(off_all); on_sorted=sorted(on_all)
for N in (3,5,10):
    print(f"bottom-{N} sum: OFF {sum(off_sorted[:N]):.4f} -> ON {sum(on_sorted[:N]):.4f}  ({sum(on_sorted[:N])-sum(off_sorted[:N]):+.4f})")
print("\n=== Top 15 regressions (ON-OFF) ===")
for r in sorted(reg,key=lambda r:r[4])[:15]:
    print(f"  {r[4]:+.4f}  {r[0][:42]:<42} p{r[1]}  {r[2]:.4f}->{r[3]:.4f}")
print("\n=== Top 15 improvements ===")
for r in sorted(imp,key=lambda r:-r[4])[:15]:
    print(f"  {r[4]:+.4f}  {r[0][:42]:<42} p{r[1]}  {r[2]:.4f}->{r[3]:.4f}")
# breakdown: regressions by gen2 jp vs other
print("\n=== regression composition ===")
print("  gen2_* regressions:",sum(1 for r in reg if r[0].startswith('gen2_')))
print("  non-gen2 regressions:",[(r[0][:30],r[1],round(r[4],3)) for r in reg if not r[0].startswith('gen2_')][:10])
