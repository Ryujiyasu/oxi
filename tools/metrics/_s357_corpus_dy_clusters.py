"""S357 — Corpus-wide low-IoU paragraph dy clustering analysis.

Discovery: the dominant Phase 2 bug pattern is NOT per-doc but a corpus-wide
+1.0pt table dy cluster (400 paragraphs across 16 docs). Per-page analysis
reveals systematic per-page baseline shifts in 0.5pt steps (0/+0.5/+1.0pt),
characteristic of accumulating rounding errors at page-top.

Usage:
  python tools/metrics/_s357_corpus_dy_clusters.py                # full corpus
  python tools/metrics/_s357_corpus_dy_clusters.py <doc_id>       # per-doc per-page
"""
import json
import sys
import collections
from pathlib import Path


def low_iou_table_dys(doc_id: str):
    p = Path(f"pipeline_data/element_iou_diff/{doc_id}.json")
    if not p.exists():
        return None
    with open(p, encoding="utf-8") as f:
        d = json.load(f)
    out = []
    for m in d.get("matches", []):
        if not m.get("matched"):
            continue
        if not m.get("in_table"):
            continue
        if (m.get("iou_raw") or 1) >= 0.9:
            continue
        dy = (m.get("oxi_y") or 0) - (m.get("word_y") or 0)
        out.append((m["word_page"], dy, m))
    return out


def corpus_clusters():
    with open("pipeline_data/element_iou_diff/_summary.json", encoding="utf-8") as f:
        summary = json.load(f)
    table_buckets = collections.Counter()
    table_docs_per_bucket = collections.defaultdict(set)
    body_buckets = collections.Counter()
    body_docs_per_bucket = collections.defaultdict(set)
    for d in summary["docs"]:
        doc_id = d["doc_id"]
        p = Path(f"pipeline_data/element_iou_diff/{doc_id}.json")
        if not p.exists():
            continue
        with open(p, encoding="utf-8") as f:
            dd = json.load(f)
        for m in dd.get("matches", []):
            if not m.get("matched"):
                continue
            if (m.get("iou_raw") or 1) >= 0.9:
                continue
            dy = (m.get("oxi_y") or 0) - (m.get("word_y") or 0)
            b = round(dy * 4) / 4
            if m.get("in_table"):
                table_buckets[b] += 1
                table_docs_per_bucket[b].add(doc_id)
            else:
                body_buckets[b] += 1
                body_docs_per_bucket[b].add(doc_id)
    print(f"Table low-IoU clusters (top 15):")
    print(f'{"dy":>8} {"count":>6} {"docs":>5}')
    for b, c in sorted(table_buckets.items(), key=lambda x: -x[1])[:15]:
        print(f"{b:>+8.2f} {c:>6} {len(table_docs_per_bucket[b]):>5}")
    print(f"\nBody low-IoU clusters (top 10):")
    print(f'{"dy":>8} {"count":>6} {"docs":>5}')
    for b, c in sorted(body_buckets.items(), key=lambda x: -x[1])[:10]:
        print(f"{b:>+8.2f} {c:>6} {len(body_docs_per_bucket[b]):>5}")


def per_doc_per_page(doc_id: str):
    dys = low_iou_table_dys(doc_id)
    if dys is None:
        print(f"No data for {doc_id}")
        return
    by_page = collections.defaultdict(list)
    for page, dy, m in dys:
        by_page[page].append(dy)
    print(f"{doc_id} table low-IoU paragraphs per page:")
    print(f'{"page":>5} {"n":>4} {"med":>7} {"top buckets":>40}')
    for page in sorted(by_page):
        ds = by_page[page]
        med = sorted(ds)[len(ds) // 2]
        c = collections.Counter(round(d * 4) / 4 for d in ds)
        top = ", ".join(f"{b:+.2f}×{n}" for b, n in c.most_common(3))
        print(f"{page:>5} {len(ds):>4} {med:>+7.2f}  {top}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        per_doc_per_page(sys.argv[1])
    else:
        corpus_clusters()
