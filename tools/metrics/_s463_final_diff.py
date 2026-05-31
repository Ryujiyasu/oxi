"""Final S463 (doc-Latin scope): ON (pipeline_data/ssim_current_post_s160.json,
rendered with OXI_S463_LATIN_BORDER=1 + doc-Latin binary) vs preserved clean OFF
(/c/tmp/s463_OFF_scores.json). Both scored vs the same Word PNGs."""
import json
OFF=json.load(open(r"C:/tmp/s463_OFF_scores.json",encoding="utf-8"))
ON=json.load(open("pipeline_data/ssim_current_post_s160.json",encoding="utf-8"))
rows=[]
for doc,pages in OFF.items():
    for pg,off in pages.items():
        on=ON.get(doc,{}).get(pg)
        if on is None: continue
        rows.append((doc,pg,off,on,on-off))
on_all=[r[3] for r in rows]; off_all=[r[2] for r in rows]
mean=lambda x:sum(x)/len(x)
print(f"pages: {len(rows)}")
print(f"OFF mean {mean(off_all):.4f}  ON mean {mean(on_all):.4f}  delta {mean(on_all)-mean(off_all):+.4f}")
buk=lambda x:f">=0.99:{sum(1 for v in x if v>=0.99)} >=0.95:{sum(1 for v in x if v>=0.95)} >=0.90:{sum(1 for v in x if v>=0.90)} <0.70:{sum(1 for v in x if v<0.70)}"
print("OFF",buk(off_all)); print("ON ",buk(on_all))
imp=[r for r in rows if r[4]>0.002]; reg=[r for r in rows if r[4]<-0.002]
print(f"improved>+0.002: {len(imp)}   regressed<-0.002: {len(reg)}   unchanged: {len(rows)-len(imp)-len(reg)}")
o=sorted(off_all); n=sorted(on_all)
for N in (3,5,10): print(f"bottom-{N}: OFF {sum(o[:N]):.4f} -> ON {sum(n[:N]):.4f} ({sum(n[:N])-sum(o[:N]):+.4f})")
print("\nregressions (<-0.002):")
for r in sorted(reg,key=lambda r:r[4])[:20]:
    print(f"  {r[4]:+.4f}  {r[0][:40]:<40} p{r[1]}  {r[2]:.4f}->{r[3]:.4f}")
print(f"\ntop improvements:")
for r in sorted(imp,key=lambda r:-r[4])[:12]:
    print(f"  {r[4]:+.4f}  {r[0][:40]:<40} p{r[1]}  {r[2]:.4f}->{r[3]:.4f}")
