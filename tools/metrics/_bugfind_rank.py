# -*- coding: utf-8 -*-
"""Rank the bug-finder JSONL: per-doc worst Libra-beats-Oxi gap (delta=libra-oxi).
Flags known-wall doc families so fresh CLEAN levers stand out.
Usage: python _bugfind_rank.py [jsonl]"""
import sys, json, collections
sys.stdout.reconfigure(encoding="utf-8")
path = sys.argv[1] if len(sys.argv) > 1 else r"c:\tmp\bugfind_s688.jsonl"
# known-wall substrings (vertical render-anchor / char-budget / form / contract)
WALLS = ["tokumei","_index-","1ec1","b35123","0e7af","683f","e3c545","gen2_","gen_",
         "tokyoshugyo","31420","15076","order_","kyodoken","de6e32","6514","a1d6","d4d126"]
rows = collections.defaultdict(list)
for ln in open(path, encoding="utf-8"):
    try: r = json.loads(ln)
    except Exception: continue
    rows[r["doc"]].append(r)
docs = []
for doc, rs in rows.items():
    worst = max(rs, key=lambda x: x["delta"])
    mean_delta = sum(x["delta"] for x in rs)/len(rs)
    docs.append((doc, worst["delta"], worst["pg"], worst["oxi"], worst["libra"], mean_delta, len(rs)))
docs.sort(key=lambda x: -x[1])
def wall(d): return any(w in d for w in WALLS)
print(f"{'doc':<46}{'worstΔ':>8}{'pg':>4}{'oxi':>7}{'libra':>7}{'meanΔ':>8}  flag")
for d,wd,pg,ox,lb,md,n in docs[:40]:
    print(f"{d[:46]:<46}{wd:>+8.3f}{pg:>4}{ox:>7.3f}{lb:>7.3f}{md:>+8.3f}  {'WALL' if wall(d) else '*** FRESH ***'}")
print(f"\n--- top FRESH (non-wall) docs ---")
fresh=[x for x in docs if not wall(x[0]) and x[1]>0.02]
for d,wd,pg,ox,lb,md,n in fresh[:20]:
    print(f"{d[:46]:<46}{wd:>+8.3f}{pg:>4}{ox:>7.3f}{lb:>7.3f}{md:>+8.3f}")
print(f"\ntotal docs scanned: {len(docs)}")
