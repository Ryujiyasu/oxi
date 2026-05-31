import json, io
base=json.load(io.open('pipeline_data/ssim_baseline.json',encoding='utf-8'))
new=json.load(io.open('pipeline_data/ssim_current_post_s160.json',encoding='utf-8'))
def flat(d): return [(doc,p,s) for doc,pg in d.items() for p,s in pg.items()]
bm={(doc,p):s for doc,p,s in flat(base)}
nm={(doc,p):s for doc,p,s in flat(new)}
common=[k for k in bm if k in nm]
bmean=sum(bm[k] for k in common)/len(common)
nmean=sum(nm[k] for k in common)/len(common)
print(f'pages compared: {len(common)}')
print(f'mean: {bmean:.4f} -> {nmean:.4f} ({nmean-bmean:+.4f})')
ups=[(k,nm[k]-bm[k]) for k in common if nm[k]-bm[k]>0.002]
dns=[(k,nm[k]-bm[k]) for k in common if nm[k]-bm[k]<-0.002]
print(f'improved: {len(ups)}  regressed: {len(dns)}')
print('--- regressions (any doc, sorted worst) ---')
for k,d in sorted(dns,key=lambda x:x[1])[:20]:
    print(f'  {d:+.4f}  {k[0][:40]} p{k[1]}')
print('--- top improvements ---')
for k,d in sorted(ups,key=lambda x:-x[1])[:12]:
    print(f'  {d:+.4f}  {k[0][:40]} p{k[1]}')
# buckets
def buck(m): 
    import collections
    c=collections.Counter()
    for k in common:
        v=m[k]
        c['<0.70']+= v<0.70; c['>=0.90']+=v>=0.90; c['>=0.95']+=v>=0.95; c['>=0.99']+=v>=0.99
    return c
print('buckets base:', dict(buck(bm)))
print('buckets new :', dict(buck(nm)))
