# -*- coding: utf-8 -*-
"""Both populations of the saved page-break marker, with Word's verdict.

`<w:lastRenderedPageBreak/>` is a CACHE of where Word broke the page when the
file was last saved, not a statement about where it breaks now. Oxi honours it
when the page is more than half consumed; the code's own note names
`policies__0353d0b2a7f98e13` as a document where that is wrong, and records that
S822's physical test (`remaining - est < K`) was rejected in 2026-07-13 because
two ENGLISH documents needed the saved breaks. Those documents no longer reach
this path at all -- S836 retired the block-level respect for Latin bodies -- so
the population that vetoed the test is gone and the question is open again.

This walks the JA corpus, reads every LRPB site's geometry from `OXI_DBG_LRPB`,
and asks Word's own pagination whether that paragraph actually starts a page.
A site Word starts a page at is LIVE; one it does not is STALE.

    python _lrpb_census.py [set ...]        sets: blind50 blindB50 (default both)
"""
import json
import os
import re
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
GDI = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
BENCH = REPO / "pipeline_data" / "ja_benchmark"
SETS = {"blind50": ("_final_jablind50.json", "p1_blind50"),
        "blindB50": ("_final_jablindB50.json", "p1_blindB50")}

LINE = re.compile(
    r"\[LRPB\] pg=(\d+) est=([\d.]+) remaining=([-\d.]+) consumed=([-\d.]+) "
    r"half=([\d.]+) slack=([-\d.]+) fires=(true|false) text=\"(.*)\"")


def norm(s):
    return "".join(s.split()).replace("\\u{3000}", "")[:14]


def main():
    want = [a for a in sys.argv[1:] if a in SETS] or list(SETS)
    rows = []
    for name in want:
        final, sub = SETS[name]
        docs = []
        for t, lst in json.load(open(BENCH / final, encoding="utf-8")).items():
            for c in lst:
                p = Path(c["path"])
                docs.append((f"{p.parent.name}__{p.stem}", str(p.resolve())))
        for did, path in docs:
            wf = BENCH / sub / "word" / f"{did}.json"
            if not wf.exists():
                continue
            # Word's own answer: which paragraphs start a page
            W = json.load(open(wf, encoding="utf-8"))["paragraphs"]
            starts, seen = set(), set()
            for r in W:
                if r["page"] not in seen:
                    seen.add(r["page"])
                    starts.add(norm(r["text"]))
            env = dict(os.environ, OXI_DBG_LRPB="1")
            try:
                out = subprocess.run([str(GDI), path, str(Path(os.environ.get("TEMP", ".")) / "lrpb")],
                                     capture_output=True, env=env, timeout=600)
            except subprocess.TimeoutExpired:
                print(f"  !! timeout {did}")
                continue
            for m in LINE.finditer(out.stderr.decode("utf8", "replace")):
                pg, est, rem, cons, half, slack, fires, text = m.groups()
                rows.append({
                    "doc": did, "pg": int(pg), "est": float(est), "remaining": float(rem),
                    "consumed": float(cons), "half": float(half), "fires": fires == "true",
                    "live": norm(text) in starts, "text": text[:18],
                })
    if not rows:
        print("no LRPB sites found")
        return
    live = [r for r in rows if r["live"]]
    stale = [r for r in rows if not r["live"]]
    print(f"LRPB sites: {len(rows)}   LIVE (Word starts a page there) {len(live)}   STALE {len(stale)}\n")
    for tag, pop in (("LIVE", live), ("STALE", stale)):
        if not pop:
            continue
        rem = sorted(r["remaining"] - r["est"] for r in pop)
        print(f"  {tag:5s} n={len(pop):3d}  remaining-est: min {rem[0]:8.2f}  "
              f"p25 {rem[len(rem)//4]:8.2f}  median {rem[len(rem)//2]:8.2f}  "
              f"p75 {rem[3*len(rem)//4]:8.2f}  max {rem[-1]:8.2f}")
    print("\n  a threshold K on (remaining - est) separates them iff LIVE stays below "
          "it and STALE above")
    for k in (10, 20, 28, 40, 60, 80, 120):
        wrong_live = sum(1 for r in live if r["remaining"] - r["est"] > k)
        wrong_stale = sum(1 for r in stale if r["remaining"] - r["est"] <= k)
        print(f"    K={k:>4}: LIVE misread {wrong_live:>3}/{len(live)}   "
              f"STALE misread {wrong_stale:>3}/{len(stale)}")
    # ★The population that matters is the FIRING one: a site Oxi does not
    # honour costs nothing either way, and a LIVE site with lots of room
    # breaks for its own reason (a hard break, a section) whether or not the
    # marker is read.
    fl = [r for r in rows if r["fires"] and r["live"]]
    fs = [r for r in rows if r["fires"] and not r["live"]]
    print(f"\n  FIRING sites: {len(fl)+len(fs)}  (LIVE {len(fl)} / STALE {len(fs)})")
    for tag, pop in (("LIVE", fl), ("STALE", fs)):
        if not pop:
            continue
        v = sorted(r["remaining"] - r["est"] for r in pop)
        print(f"    {tag:5s} n={len(pop):3d}  remaining-est: min {v[0]:8.2f} "
              f"median {v[len(v)//2]:8.2f}  max {v[-1]:8.2f}")
    print("    threshold on the FIRING population only:")
    for k in (5, 10, 15, 20, 28, 40, 60):
        wl = sum(1 for r in fl if r["remaining"] - r["est"] > k)
        ws = sum(1 for r in fs if r["remaining"] - r["est"] <= k)
        print(f"      K={k:>3}: LIVE lost {wl:>3}/{len(fl)}   STALE still fires {ws:>3}/{len(fs)}")

    # Is staleness a DOCUMENT property? S811 already distrusts the markers
    # per document (Latin bodies whose fonts were substituted); if a file
    # whose saved layout has moved has ALL of its markers stale, the same
    # shape of test works here without needing a per-site threshold.
    per = {}
    for r in rows:
        d = per.setdefault(r["doc"], [0, 0])
        d[0 if r["live"] else 1] += 1
    mixed = {d: v for d, v in per.items() if v[0] and v[1]}
    pure_stale = {d: v for d, v in per.items() if v[1] and not v[0]}
    print(f"\n  documents with LRPB sites: {len(per)}")
    print(f"    all-LIVE      {sum(1 for v in per.values() if v[0] and not v[1]):>3}")
    print(f"    all-STALE     {len(pure_stale):>3}   {list(pure_stale)}")
    print(f"    MIXED         {len(mixed):>3}   (a per-document test cannot help these)")
    for d, v in sorted(mixed.items(), key=lambda kv: -kv[1][1]):
        print(f"      {d:<42} live {v[0]:>3}  stale {v[1]:>3}")

    print("\n  STALE sites that currently FIRE (the ones costing pages):")
    for r in sorted((r for r in stale if r["fires"]), key=lambda r: -(r["remaining"] - r["est"]))[:12]:
        print(f"    {r['doc']:<40} pg{r['pg']:>3} rem-est {r['remaining']-r['est']:8.2f}  {r['text']!r}")


if __name__ == "__main__":
    main()
