"""S387 — per-page baseline STEP diagnostic + corpus ceiling.

KEY INSIGHT (S387): the Phase 2 gate metric (element_iou_diff.py) subtracts
a DOC-WIDE per-cohort (body/table) median dy before computing iou_pos
(see element_iou_diff.py:334-365). Therefore a UNIFORM per-doc offset is
FULLY ABSORBED and costs zero IoU. The +1.0pt absolute cluster (S357),
chased S357-S386, is largely uniform-per-page and thus mostly absorbed.

The REAL lever is WITHIN-DOC variance: per-PAGE baseline STEPS (a page's
content uniformly offset from the doc-wide median) and within-page spread.

This tool:
  1. Per doc, computes each (page, in_table) cohort's median dy and the
     spread across pages → ranks docs by per-page step magnitude.
  2. Simulates eliminating the per-page step (use a per-PAGE per-cohort
     median instead of doc-wide) and reports the corpus mean_iou CEILING.

Run from repo root:
    python tools/metrics/_s387_per_page_step.py            # corpus ceiling + ranking
    python tools/metrics/_s387_per_page_step.py <doc8>     # one doc detail

Reads pipeline_data/element_iou_diff/*.json (run element_iou_diff.py first).
"""
from __future__ import annotations
import json, glob, os, sys
from collections import defaultdict

IOU_DIR = os.path.join(os.path.dirname(__file__), "..", "..",
                       "pipeline_data", "element_iou_diff")


def med(xs):
    xs = sorted(xs)
    return xs[len(xs) // 2] if xs else 0.0


def pos_iou(wy, wh, oy, oh):
    h = max(wh, oh, 1e-9)
    return max(0.0, 1.0 - abs(oy - wy) / h)


def doc_mean(matches, per_page):
    g = defaultdict(list)
    key = ((lambda x: (x["oxi_page"], x["in_table"])) if per_page
           else (lambda x: x["in_table"]))
    for x in matches:
        g[key(x)].append(x["oxi_y"] - x["word_y"])
    offs = {k: med(v) for k, v in g.items()}
    ious = [pos_iou(x["word_y"], x["word_h"],
                    x["oxi_y"] - offs[key(x)], x["oxi_h"]) for x in matches]
    nhigh = sum(1 for i in ious if i >= 0.99)
    return (sum(ious) / len(ious) if ious else 0.0), nhigh, len(ious)


def detail(doc8):
    p = glob.glob(os.path.join(IOU_DIR, doc8 + "*.json"))
    if not p:
        print("no doc", doc8); return
    d = json.load(open(p[0], encoding="utf-8"))
    m = [x for x in d["matches"] if x["matched"]]
    print("doc", doc8, "mean_iou", d["mean_iou"],
          "median_dy_table", d.get("median_dy_table"),
          "median_dy_body", d.get("median_dy_body"))
    byp = defaultdict(list)
    for x in m:
        byp[(x["oxi_page"], x["in_table"])].append(x["oxi_y"] - x["word_y"])
    for k in sorted(byp):
        v = sorted(byp[k]); n = len(v)
        print("  page %2d %s n=%3d med=%+.2f range[%+.2f,%+.2f]" % (
            k[0], "tab" if k[1] else "bod", n, v[n // 2], min(v), max(v)))
    cur = doc_mean(m, False); pp = doc_mean(m, True)
    print("  doc-wide median (gate):  mean=%.4f n_high=%d/%d" % cur)
    print("  per-page median (ceil):  mean=%.4f n_high=%d/%d" % pp)


def corpus():
    cur, pp, rows = [], [], []
    for p in glob.glob(os.path.join(IOU_DIR, "*.json")):
        if os.path.basename(p).startswith("_"):
            continue
        try:
            d = json.load(open(p, encoding="utf-8"))
        except Exception:
            continue
        m = [x for x in d.get("matches", []) if x.get("matched")]
        if not m or "mean_iou" not in d:
            continue
        a, _, _ = doc_mean(m, False)
        b, _, _ = doc_mean(m, True)
        cur.append(a); pp.append(b)
        rows.append((b - a, os.path.basename(p)[:8], a, b))
    n = len(cur)
    print("docs", n)
    print("current corpus mean      : %.4f" % (sum(cur) / n))
    print("per-page-step-eliminated : %.4f" % (sum(pp) / n))
    print("CEILING gain             : +%.4f" % ((sum(pp) - sum(cur)) / n))
    print("\ntop per-doc upside from eliminating per-page step:")
    rows.sort(reverse=True)
    for g, did, a, b in rows[:15]:
        print("  %s  %.4f -> %.4f  (+%.4f)" % (did, a, b, g))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        detail(sys.argv[1])
    else:
        corpus()
