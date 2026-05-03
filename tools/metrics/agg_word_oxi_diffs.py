"""Aggregate divergence patterns across docs to identify common cause clusters."""
import json, os, glob, sys
from collections import Counter
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DIR = "C:/Users/ryuji/AppData/Local/Temp/diffv2_out"
files = sorted(glob.glob(os.path.join(DIR, "*.json")))
print(f"Loaded {len(files)} doc reports")

# For each doc, find the FIRST big-jump paragraph (where Δy goes from <3 to >5)
# That's the divergence epicenter
print(f"\n=== Per-doc divergence epicenter ===")
for f in files:
    with open(f, encoding='utf-8') as fh:
        d = json.load(fh)
    doc_id = d["doc_id"][:40]
    matches = d["matches"]
    prev_y = 0
    found = False
    for m in matches:
        dy = m["divergence"]["y_diff"]
        if dy is None: continue
        if abs(dy) > 5 and abs(prev_y) <= 3:
            wp = m["word"]
            print(f"  {doc_id:40s}  P{wp['idx']:>3} '{wp['text'][:30]:30s}' Δy=+{dy:>5.2f} (prev was {prev_y:+.2f})")
            found = True
            break
        prev_y = dy
    if not found:
        print(f"  {doc_id:40s}  no clean epicenter")

# Per-doc: what Y diff value is most common in the "drifting" range?
print(f"\n=== Drift mode (most common Δy bucket per doc) ===")
for f in files:
    with open(f, encoding='utf-8') as fh:
        d = json.load(fh)
    doc_id = d["doc_id"][:40]
    matches = d["matches"]
    bucket = Counter()
    for m in matches:
        dy = m["divergence"]["y_diff"]
        if dy is None: continue
        b = round(dy / 5) * 5  # bucket by 5pt
        bucket[b] += 1
    top = bucket.most_common(3)
    print(f"  {doc_id:40s}  top buckets: {top}")

# Common patterns across all docs
print(f"\n=== Cross-doc Δy histogram (5pt buckets) ===")
all_dy_buckets = Counter()
for f in files:
    with open(f, encoding='utf-8') as fh:
        d = json.load(fh)
    for m in d["matches"]:
        dy = m["divergence"]["y_diff"]
        if dy is None: continue
        b = round(dy / 5) * 5
        all_dy_buckets[b] += 1
for b in sorted(all_dy_buckets.keys()):
    cnt = all_dy_buckets[b]
    bar = "#" * min(cnt, 50)
    print(f"  Δy={b:>+4}pt: {cnt:>4d} {bar}")
