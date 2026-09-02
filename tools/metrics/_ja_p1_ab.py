# -*- coding: utf-8 -*-
"""Full paragraph-level Phase-1 A/B over the JA blind sets, one flag toggled.

`_jp_p1_flag_census.py` answers the page-COUNT question; this answers the
per-paragraph one (pass/fail, score, delta histogram), which is the actual
Phase-1 gate. Both arms are rendered with the SAME binary in the same run and
NOTHING is read from cache — `phase_oxi`-style "skip if the JSON exists"
scores whatever binary wrote it last (2026-09-01 incident).

    python _ja_p1_ab.py OXI_S1290_DISABLE [set ...] [--only substr]

Arm A = flag set (ship OFF), arm B = unset (ship ON, the default).
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

FLAG = sys.argv[1]
args = sys.argv[2:]
only = None
if "--only" in args:
    only = args[args.index("--only") + 1]
    args = [a for a in args if a != "--only" and a != only]
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
        for did, path, outdir in docs(setname):
            if only and only not in did:
                continue
            wf = outdir / "word" / f"{did}.json"
            if not wf.exists():
                continue
            word = json.loads(wf.read_text(encoding="utf-8"))
            arms = {}
            for arm, flag_on in (("off", True), ("on", False)):
                if flag_on:
                    os.environ[FLAG] = "1"
                else:
                    os.environ.pop(FLAG, None)
                try:
                    arms[arm] = PD.diff_doc(did, word, MO.measure_doc(path))
                except Exception as e:
                    print(f"  {did}: {arm} FAIL {str(e)[:70]}")
                    arms[arm] = None
            os.environ.pop(FLAG, None)
            if not arms["on"] or not arms["off"]:
                continue
            rows.append((setname, did, arms["on"], arms["off"]))
            a, b = arms["on"], arms["off"]
            if a["pass"] != b["pass"] or abs(a["score"] - b["score"]) > 1e-4 \
                    or a["page_count_delta"] != b["page_count_delta"]:
                verdict = ("BETTER" if (a["pass"], round(a["score"], 4))
                           > (b["pass"], round(b["score"], 4)) else "WORSE")
                print(f"  {setname:<9} {did:<34} ON pass={a['pass']} "
                      f"score={a['score']:.4f} pcd={a['page_count_delta']:+d} | "
                      f"OFF pass={b['pass']} score={b['score']:.4f} "
                      f"pcd={b['page_count_delta']:+d}  {verdict}")
                print(f"      hist ON  {a['delta_histogram']}")
                print(f"      hist OFF {b['delta_histogram']}")

    n = len(rows)
    if not n:
        print("no docs measured")
        return
    on_pass = sum(1 for r in rows if r[2]["pass"])
    off_pass = sum(1 for r in rows if r[3]["pass"])
    on_score = sum(r[2]["score"] for r in rows) / n
    off_score = sum(r[3]["score"] for r in rows) / n
    on_pcd = sum(abs(r[2]["page_count_delta"]) for r in rows)
    off_pcd = sum(abs(r[3]["page_count_delta"]) for r in rows)
    flips = [r for r in rows if r[2]["pass"] != r[3]["pass"]]
    print(f"\n=== {FLAG} over {n} docs ===")
    print(f"PASS      ship ON {on_pass}/{n}   ship OFF {off_pass}/{n}")
    print(f"mean score ship ON {on_score:.4f}   ship OFF {off_score:.4f}")
    print(f"sum|pcd|  ship ON {on_pcd}   ship OFF {off_pcd}")
    for r in flips:
        print(f"  FLIP {r[1]}: ON pass={r[2]['pass']} OFF pass={r[3]['pass']}")


if __name__ == "__main__":
    main()
