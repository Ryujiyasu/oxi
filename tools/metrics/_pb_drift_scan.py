# -*- coding: utf-8 -*-
"""Classify a page-break failure: one missing term, or a knife-edge?

Both look the same in the pagination gate (delta = -1 on one paragraph), but
they need opposite work. Walking the page that over-packed and printing how the
Word-vs-Oxi position difference STEPS between consecutive paragraphs separates
them in one read:

  a single large step   -> a term is missing at that paragraph (S1249 was 35.70pt
                           at one heading; fixable, and usually shared by a class)
  steps all ~0, the
  last line just fits   -> the page bottom is a knife-edge (sub-pt calibration)

Word's y is Information(6) (the cursor before the paragraph) and Oxi's is the
line-box top, so only the STEP of the difference is meaningful, never its value.

    python tools/metrics/_pb_drift_scan.py              # every failing doc
    python tools/metrics/_pb_drift_scan.py <doc_id> ...
"""
import json, os, re, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pagination_diff as PD

BENCH = "pipeline_data/en_benchmark"
SETS = {"blind50": "_final_blind50.json", "blindB50": "_final_blindB50.json",
        "blindC50": "_final_blindC50.json", "next50": "_final_next50.json"}
norm = lambda s: re.sub(r"\s+", "", s)[:22]


def load(setname, did):
    b = "%s/p1_%s" % (BENCH, setname)
    w = json.load(open("%s/word/%s.json" % (b, did), encoding="utf-8"))["paragraphs"]
    o = json.load(open("%s/oxi/%s.json" % (b, did), encoding="utf-8"))["pages"]
    return w, o


def scan(setname, did, want=None):
    w, o = load(setname, did)
    r = PD.diff_doc(did, {"paragraphs": w},
                    json.load(open("%s/p1_%s/oxi/%s.json" % (BENCH, setname, did),
                                   encoding="utf-8")))
    sites = [m for m in r["matches"] if m.get("page_delta")]
    if not sites:
        return
    print("=== %-34s %-9s score %.4f  %d sites ===" % (did, setname, r["score"], len(sites)))
    for site in sites[:3]:
        pg = site["oxi_page"]          # the page Oxi over-packed (or under-filled)
        wrecs = [x for x in w if x["page"] == pg]
        orecs = [x for x in o.get(str(pg), []) if x.get("para_idx") is not None]
        if not wrecs or not orecs:
            continue
        oi, prev, worst = 0, None, (0.0, "")
        steps = []
        for x in wrecs:
            key = norm(x["text"])
            j = oi
            while j < len(orecs) and norm(orecs[j]["text"]) != key:
                j += 1
            if j >= len(orecs):
                continue
            oi = j + 1
            if orecs[j].get("y") is None or x.get("y") is None:
                continue          # a record without a position says nothing here
            d = orecs[j]["y"] - x["y"]
            if prev is not None:
                step = d - prev
                steps.append(step)
                if abs(step) > abs(worst[0]):
                    worst = (step, key)
            prev = d
        span = (max(steps) - min(steps)) if steps else 0.0
        print("   site d=%+d page %d: %d shared paras | worst step %+.2f at %r | span %.2f"
              % (site["page_delta"], pg, len(steps) + 1, worst[0], worst[1], span))


if __name__ == "__main__":
    want = set(sys.argv[1:])
    for s, sel in SETS.items():
        final = json.load(open(os.path.join(BENCH, sel), encoding="utf-8"))
        for t, lst in final.items():
            for c in lst:
                did = "%s__%s" % (os.path.basename(os.path.dirname(c["path"])),
                                  os.path.splitext(os.path.basename(c["path"]))[0])
                if want and did not in want:
                    continue
                wf = "%s/p1_%s/word/%s.json" % (BENCH, s, did)
                of = "%s/p1_%s/oxi/%s.json" % (BENCH, s, did)
                if not (os.path.exists(wf) and os.path.exists(of)):
                    continue
                rr = PD.diff_doc(did, json.load(open(wf, encoding="utf-8")),
                                 json.load(open(of, encoding="utf-8")))
                if rr["pass"]:
                    continue
                try:
                    scan(s, did)
                except Exception as e:
                    print("=== %s (%s) SCAN FAILED: %s" % (did, s, str(e)[:60]))
