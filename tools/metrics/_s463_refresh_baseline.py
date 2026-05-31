"""Ratchet ssim_baseline.json UP for the S463-improved pages only (ON > OFF).
ON = pipeline_data/ssim_current_post_s160.json (default-ON render),
OFF = /c/tmp/s463_OFF_scores.json. Unchanged pages keep their baseline value."""
import json
BASE="pipeline_data/ssim_baseline.json"
ON=json.load(open("pipeline_data/ssim_current_post_s160.json",encoding="utf-8"))
OFF=json.load(open(r"C:/tmp/s463_OFF_scores.json",encoding="utf-8"))
base=json.load(open(BASE,encoding="utf-8"))
changed=[]
for doc,pages in ON.items():
    for pg,on in pages.items():
        off=OFF.get(doc,{}).get(pg)
        if off is None: continue
        if on-off>0.002:  # real S463 improvement
            old=base.get(doc,{}).get(pg)
            base.setdefault(doc,{})[pg]=on
            changed.append((doc,pg,old,on))
json.dump(base,open(BASE,"w",encoding="utf-8"),indent=2,ensure_ascii=False)
a=[v for p in base.values() for v in p.values()]
print(f"updated {len(changed)} pages; new baseline mean {sum(a)/len(a):.4f}")
for d,p,o,n in sorted(changed,key=lambda x:-(x[3]-(x[2] or 0)))[:8]:
    print(f"  {d[:38]:<38} p{p}  {o}->{round(n,4)}")
