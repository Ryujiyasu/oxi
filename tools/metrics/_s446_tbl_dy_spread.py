"""S446 (2026-05-30): gate-relevant table-dy diagnostic + docGrid-type discriminator test.

NOTE: Finding 2 below was RETRACTED in S447 (same-session correction). The
original claim — "a uniform per-row height delta is median-cancelled / gate-
invisible" — is FALSE. A uniform per-ROW HEIGHT delta accumulates into a linear
POSITION ramp; the median only removes the ramp's midpoint, so the ramp's spread
(residual) IS visible to the gate. (What median DOES cancel is a uniform per-
element POSITION offset — the whole table shifted down by a constant — which is a
different thing.) Verified on 7ead: dy ramps +0.10 -> -9.65 monotonically over 8
rows (~1.2pt/row under-height), residual ±5pt = gate-visible. The real reason
S183/S208/S223/S445 failed is MIS-TARGETING (the gate condition fired on docs
without the overhead, injecting a false ramp there), NOT median-cancellation.

Finding 1 (docGrid type) stands. Finding 2 is retracted; see S447 memory.

ORIGINAL TEXT (kept for the record, Finding 2 is wrong):

1. docGrid TYPE does NOT discriminate Oxi-short vs Oxi-tall table drift.
   Both `lines` and `linesAndChars` docs span both dy directions
   (a47e lines=+2.2 vs 7ead lines=-5.8; b35 linesAndChars=+1.4 vs
   d4d126 linesAndChars=-1.2). So docGrid type is NOT the missing
   discriminator behind the S183/S208/S223/S445 falsifications.

2. [RETRACTED] The Phase-2 gate (element_iou_diff.py:360-366) subtracts the
   per-(in_table) MEDIAN dy before computing iou_pos. [The original inference
   that this makes a uniform per-row HEIGHT delta gate-invisible is wrong — see
   the NOTE above. A height delta forms a ramp, which the gate sees.]

Ranking docs by gate-relevant residual spread (sd of dy after median removal)
confirms the largest spreads are all matcher/COM artifacts (3a4f sd=50 multi-col
kinsoku, 31420a sd=16.5 multi-col form S429) or Phase-1-blocked (ed025 sd=14.9).
Per-row SHAPE is doc-specific (7ead = clean under-height ramp ~1.2pt/row;
a47e = one over-tall row +4.2 then gradual under-height), reconfirming no
convergent rule. Each fixable doc has a single small table -> ~+0.001 corpus
per doc at high mis-targeting risk (Phase-3 stop). No clean convergent Phase-2
layout target remains (reconfirms S433 from the gate's own viewpoint).

Run from repo root:
  python tools/metrics/_s446_tbl_dy_spread.py            # spread ranking
  python tools/metrics/_s446_tbl_dy_spread.py --docgrid  # docGrid-type vs dy
"""
from __future__ import annotations
import glob
import io
import json
import os
import re
import statistics as st
import sys
import zipfile

EIOU = "pipeline_data/element_iou_diff"
DOCX = "tools/golden-test/documents/docx"


def _find_list(o):
    if isinstance(o, list) and o and isinstance(o[0], dict):
        return o
    if isinstance(o, dict):
        for v in o.values():
            r = _find_list(v)
            if r:
                return r
    return None


def tbl_dys(did):
    f = f"{EIOU}/{did}.json"
    if not os.path.exists(f):
        return []
    d = json.load(io.open(f, encoding="utf-8"))
    lst = _find_list(d) or []
    return [
        (e["oxi_y"] - e["word_y"], e.get("iou_pos"))
        for e in lst
        if e.get("in_table")
        and e.get("oxi_y") is not None
        and e.get("word_y") is not None
    ]


def docgrid(p):
    try:
        x = zipfile.ZipFile(p).read("word/document.xml").decode("utf-8", "replace")
    except Exception:
        return None
    m = re.search(r'<w:docGrid w:type="([^"]+)"', x)
    cs = re.search(r'<w:docGrid[^>]*w:charSpace="(-?\d+)"', x)
    return (m.group(1) if m else "none", cs.group(1) if cs else "")


def rank_spread():
    res = []
    for f in glob.glob(f"{EIOU}/*.json"):
        did = os.path.basename(f)[:-5]
        if did.startswith("_"):
            continue
        dys = tbl_dys(did)
        if len(dys) < 8:
            continue
        vals = [x[0] for x in dys]
        med = st.median(vals)
        sd = st.pstdev([v - med for v in vals])
        iou = st.mean([x[1] for x in dys if x[1] is not None])
        res.append((sd, did, len(vals), med, iou))
    res.sort(reverse=True)
    print(f"{'doc':8}{'n':>5}{'sd_resid':>10}{'median':>9}{'tbl_iou':>9}")
    for sd, did, n, med, iou in res:
        print(f"{did[:6]:8}{n:>5}{sd:>10.2f}{med:>+9.2f}{iou:>9.3f}")


def by_docgrid():
    from collections import defaultdict

    g = defaultdict(list)
    for p in glob.glob(f"{DOCX}/*.docx"):
        dg = docgrid(p)
        if dg is None:
            continue
        did = os.path.basename(p)[:12]
        dys = tbl_dys(did)
        if len(dys) < 4:
            continue
        med = st.median([x[0] for x in dys])
        g[dg[0]].append((med, did, len(dys), dg[1]))
    print("=== table median_dy grouped by docGrid type ===")
    for typ, vals in sorted(g.items()):
        mds = [v[0] for v in vals]
        print(f"\n[{typ}] n_docs={len(vals)} median_of_medians={st.median(mds):+.2f}")
        for md, did, n, cs in sorted(vals):
            print(f"   {did[:6]} dy={md:+6.2f} n={n:3d} charSpace={cs}")


if __name__ == "__main__":
    if "--docgrid" in sys.argv:
        by_docgrid()
    else:
        rank_spread()
