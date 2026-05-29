"""S445: per-doc compare current element_iou + pagination summaries vs the
backed-up baseline (C:/tmp/s445/). Shows only docs that MOVED."""
import json, io, sys
sys.stdout.reconfigure(encoding="utf-8")

def load(p):
    return json.load(io.open(p, encoding="utf-8"))

base = load(r"C:/tmp/s445/baseline_eiou.json")
cur = load(r"c:/Users/ryuji/oxi-main/pipeline_data/element_iou_diff/_summary.json")
bd = {d["doc_id"]: d["mean_iou"] for d in base["docs"]}
cd = {d["doc_id"]: d["mean_iou"] for d in cur["docs"]}
print(f"IOU mean: baseline {base['mean_iou']:.4f} -> current {cur['mean_iou']:.4f}  (Δ {cur['mean_iou']-base['mean_iou']:+.4f})")
print(f"IOU pass: baseline {base['n_pass']} -> current {cur['n_pass']}")
print("--- per-doc IoU moves (|Δ|>=0.001) ---")
moved = []
for doc in sorted(set(bd) | set(cd)):
    b = bd.get(doc); c = cd.get(doc)
    if b is None or c is None:
        print(f"  {doc}: MISSING base={b} cur={c}"); continue
    if abs(c - b) >= 0.001:
        moved.append((doc, b, c))
for doc, b, c in sorted(moved, key=lambda x: x[2] - x[1]):
    tag = "UP" if c > b else "DOWN"
    print(f"  {doc}: {b:.4f} -> {c:.4f}  ({c-b:+.4f}) {tag}")
print(f"  ({len(moved)} docs moved; {sum(1 for _,b,c in moved if c>b)} up, {sum(1 for _,b,c in moved if c<b)} down)")

# pagination
try:
    pb = load(r"C:/tmp/s445/baseline_pag.json")
    pc = load(r"c:/Users/ryuji/oxi-main/pipeline_data/pagination_diff/_summary.json")
    print(f"\nPAGINATION: baseline n_pass {pb['n_pass']}/{pb['n_total']} -> current {pc['n_pass']}/{pc['n_total']}")
    pbd = {d['doc_id']: d.get('pass') for d in pb.get('docs', [])}
    pcd = {d['doc_id']: d.get('pass') for d in pc.get('docs', [])}
    for doc in sorted(set(pbd) | set(pcd)):
        if pbd.get(doc) != pcd.get(doc):
            print(f"  PAG CHANGE {doc}: {pbd.get(doc)} -> {pcd.get(doc)}")
except Exception as e:
    print("pagination compare skipped:", e)
