# -*- coding: utf-8 -*-
"""Single-arm page_count_delta census over the JP corpus (every doc with
cached Word truth under pipeline_data/pagination_word/)."""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import measure_pagination_oxi as MO  # noqa: E402
import pagination_diff as PD  # noqa: E402

REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
WORD_DIR = os.path.join(REPO, "pipeline_data", "pagination_word")
DOCS_DIR = os.path.join(REPO, "tools", "golden-test", "documents", "docx")

by_id = {}
for f in sorted(os.listdir(DOCS_DIR)):
    if f.endswith(".docx"):
        by_id[MO.doc_id_from_filename(f)] = os.path.join(DOCS_DIR, f)

rows = []
for wf in sorted(os.listdir(WORD_DIR)):
    if not wf.endswith(".json") or wf.startswith("_"):
        continue
    did = wf[:-5]
    path = by_id.get(did)
    if not path:
        continue
    word = json.load(open(os.path.join(WORD_DIR, wf), encoding="utf-8"))
    try:
        oxi = MO.measure_doc(path)
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR {did}: {exc}")
        continue
    d = PD.diff_doc(did, word, oxi)
    rows.append({
        "doc": did, "pcd": d["page_count_delta"], "pass": d["pass"],
        "score": d["score"], "word_pages": d.get("word_n_pages"),
        "oxi_pages": d.get("oxi_n_pages"),
    })

json.dump(rows, open(os.path.join(REPO, "pipeline_data", "en_benchmark",
                                  "_jp_pcd_census_result.json"), "w", encoding="utf-8"))
bad = [r for r in rows if r["pcd"] != 0]
bad.sort(key=lambda r: (-abs(r["pcd"]), r["score"]))
print(f"\n{len(rows)} docs; pcd!=0: {len(bad)}")
for r in bad:
    print(f"  pcd={r['pcd']:+d}  W{r['word_pages']}/O{r['oxi_pages']}"
          f"  {'PASS' if r['pass'] else 'FAIL'} {r['score']:.4f}  {r['doc']}")
