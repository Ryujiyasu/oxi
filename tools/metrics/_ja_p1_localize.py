# -*- coding: utf-8 -*-
"""Where does a doc's pagination FIRST diverge, and what is on that boundary?

A cascade reports thousands of wrong paragraphs but has one cause: the first
page whose content stops matching. This prints, per page, how many matched
paragraphs carry each delta, then the last few Word paragraphs of the last
clean page and the first few of the page after -- which is where the missing
(or extra) break lives.

    python _ja_p1_localize.py <set> <doc-substring>
"""
import json
import os
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
REPO = Path(__file__).resolve().parents[2]
BENCH = REPO / "pipeline_data" / "ja_benchmark"
SETS = {"blind50": ("_final_jablind50.json", "p1_blind50"),
        "blindB50": ("_final_jablindB50.json", "p1_blindB50")}


def dec(t):
    try:
        return t.encode("latin-1").decode("cp932")
    except Exception:
        return t


def main():
    setname, needle = sys.argv[1], sys.argv[2]
    import measure_pagination_oxi as MO
    import pagination_diff as PD

    manifest, outdir = SETS[setname]
    data = json.loads((BENCH / manifest).read_text(encoding="utf-8"))
    target = None
    for _t, lst in data.items():
        for c in lst:
            p = Path(c["path"])
            did = f"{p.parent.name}__{p.stem}"
            if needle in did:
                target = (did, str(p.resolve()), BENCH / outdir)
    if not target:
        print("no doc matched")
        return
    did, path, od = target
    word = json.loads((od / "word" / f"{did}.json").read_text(encoding="utf-8"))
    r = PD.diff_doc(did, word, MO.measure_doc(path))
    print(f"{did}: word {r['word_n_pages']}pg  oxi {r['oxi_n_pages']}pg  "
          f"pcd {r['page_count_delta']:+d}  score {r['score']}")
    print(f"  word blank pages {r['word_blank_pages']}  oxi blank {r['oxi_blank_pages']}")

    per = defaultdict(Counter)
    for m in r["matches"]:
        # An unmatched paragraph carries page_delta None; keep it visible as a
        # bucket of its own rather than letting it crash the sort.
        d = m["page_delta"]
        per[m["word_page"]][d if d is not None else "unmatched"] += 1
    first_bad = None
    print(f"\n  {'wordpg':>6}  deltas")
    for pg in sorted(per):
        c = dict(sorted(per[pg].items(), key=lambda kv: (isinstance(kv[0], str), kv[0])))
        flag = ""
        if first_bad is None and any(k != 0 for k in c):
            first_bad = pg
            flag = "   <<< FIRST DIVERGENCE"
        print(f"  {pg:>6}  {c}{flag}")

    if first_bad:
        print(f"\n  --- Word page {first_bad-1} (last clean), tail ---")
        prev = [p for p in word["paragraphs"] if p["page"] == first_bad - 1]
        for p in prev[-6:]:
            print(f"    i={p['i']:>4} y={p['y']:>7.2f} x={p['x']:>7.2f} "
                  f"tbl={p.get('in_table')} {dec(p['text'])[:40]!r}")
        print(f"  --- Word page {first_bad} (first diverging), head ---")
        cur = [p for p in word["paragraphs"] if p["page"] == first_bad]
        for p in cur[:6]:
            print(f"    i={p['i']:>4} y={p['y']:>7.2f} x={p['x']:>7.2f} "
                  f"tbl={p.get('in_table')} {dec(p['text'])[:40]!r}")


if __name__ == "__main__":
    main()
