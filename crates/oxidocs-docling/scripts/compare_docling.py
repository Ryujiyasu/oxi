"""Compare an export folder with docling's tests/data/docx/groundtruth: share of docling's body texts present.

    python compare_docling.py <export dir> <docling/tests/data/docx/groundtruth>
"""
import json, glob, os, re, sys
S, GT = sys.argv[1], sys.argv[2]
norm = lambda s: re.sub(r"\s+", "", s)
def texts(d, body_only=True):
    out = [t["text"] for t in d["texts"] if not body_only or t.get("content_layer", "body") == "body"]
    for t in d["tables"]:
        out += [c["text"] for c in t["data"]["table_cells"]]
    return [norm(x) for x in out if norm(x)]
rows = []; tot_gt = tot_hit = 0; prov_items = items = unpl = 0
err = open(os.path.join(S, "_stderr.txt")).read() if os.path.exists(os.path.join(S, "_stderr.txt")) else ""
for f in sorted(glob.glob(os.path.join(S, "*.json"))):
    name = os.path.basename(f)[:-5]
    g = os.path.join(GT, name + ".docx.json")
    o = json.load(open(f))
    allitems = o["texts"] + o["tables"] + o["pictures"]
    items += len(allitems); prov_items += sum(1 for x in allitems if x["prov"])
    if not os.path.exists(g):
        continue
    gt = json.load(open(g))
    blob = "".join(texts(o))
    want = texts(gt)
    hit = sum(1 for w in want if w in blob)
    tot_gt += len(want); tot_hit += hit
    rows.append((hit / len(want) if want else 1.0, name, hit, len(want)))
for r, n, h, w in sorted(rows)[:8]:
    print(f"{r:6.1%} {h:4}/{w:<4} {n}")
print(f"docs compared: {len(rows)}; Docling body texts found in Oxi export: {tot_hit}/{tot_gt} = {tot_hit/tot_gt:.1%}")
print(f"items with prov: {prov_items}/{items}")
print("unplaced fragments total:", sum(int(m) for m in re.findall(r"(\d+) unplaced", err)))
