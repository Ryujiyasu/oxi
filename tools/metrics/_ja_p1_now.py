# -*- coding: utf-8 -*-
"""Current JA Phase-1 state over the blind sets, ranked worst-first.

The stored `_result.json` beside each set is whatever the last full run wrote
and goes stale the moment anything ships; this re-renders every doc with the
CURRENT binary and reports where the gate actually stands. Word-side JSONs are
reused (they are ground truth and do not move).

    python _ja_p1_now.py [set ...] [--save]
"""
import json
import os
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
REPO = Path(__file__).resolve().parents[2]
BENCH = REPO / "pipeline_data" / "ja_benchmark"
SETS = {"blind50": ("_final_jablind50.json", "p1_blind50"),
        "blindB50": ("_final_jablindB50.json", "p1_blindB50")}

args = sys.argv[1:]
save = "--save" in args
want = [a for a in args if a in SETS] or list(SETS)


def docs(setname):
    manifest, outdir = SETS[setname]
    data = json.loads((BENCH / manifest).read_text(encoding="utf-8"))
    for _t, lst in data.items():
        for c in lst:
            p = Path(c["path"])
            yield f"{p.parent.name}__{p.stem}", str(p.resolve()), BENCH / outdir


def main():
    import measure_pagination_oxi as MO
    import pagination_diff as PD

    rows = []
    for setname in want:
        out = []
        for did, path, outdir in docs(setname):
            wf = outdir / "word" / f"{did}.json"
            if not wf.exists():
                continue
            word = json.loads(wf.read_text(encoding="utf-8"))
            try:
                r = PD.diff_doc(did, word, MO.measure_doc(path))
            except Exception as e:
                print(f"  {did}: FAIL {str(e)[:70]}")
                continue
            rec = {"doc": did, "pass": r["pass"], "score": round(r["score"], 4),
                   "pcd": r["page_count_delta"], "n": r["n_matched"],
                   "match_rate": round(r["match_rate"], 4),
                   "hist": r["delta_histogram"]}
            out.append(rec)
            rows.append((setname, rec))
        if save:
            (BENCH / SETS[setname][1] / "_result.json").write_text(
                json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")

    n = len(rows)
    npass = sum(1 for _, r in rows if r["pass"])
    print(f"\n=== JA Phase-1 now: PASS {npass}/{n} ({100*npass/max(n,1):.0f}%)  "
          f"mean score {sum(r['score'] for _, r in rows)/max(n,1):.4f}  "
          f"sum|pcd| {sum(abs(r['pcd']) for _, r in rows)}")
    fails = [(s, r) for s, r in rows if not r["pass"]]
    fails.sort(key=lambda x: (x[1]["score"], -abs(x[1]["pcd"])))
    print(f"\n{'set':<9} {'doc':<34} {'score':>7} {'pcd':>4} {'n':>5}  hist")
    for s, r in fails:
        print(f"{s:<9} {r['doc']:<34} {r['score']:>7.4f} {r['pcd']:>+4d} "
              f"{r['n']:>5}  {r['hist']}")


if __name__ == "__main__":
    main()
