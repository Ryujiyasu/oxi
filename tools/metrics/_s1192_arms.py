# -*- coding: utf-8 -*-
"""Run one golden doc through several env arms and print Phase-1 pass/score."""
import json, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import measure_pagination_oxi as MO
import pagination_diff as PD

REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
WORD_DIR = os.path.join(REPO, "pipeline_data", "pagination_word")
DOCS_DIR = os.path.join(REPO, "tools", "golden-test", "documents", "docx")

did = sys.argv[1]
arms = [a for a in sys.argv[2:]]

by_id = {f[:-5]: os.path.join(DOCS_DIR, f) for f in os.listdir(DOCS_DIR)
         if f.endswith(".docx") and not f.startswith("~$")}
path = by_id.get(did) or next(v for k, v in by_id.items() if k.startswith(did.split("_")[0]))
word = json.load(open(os.path.join(WORD_DIR, did + ".json"), encoding="utf-8"))

ALL = sorted({f for a in arms for f in a.split(",") if f})
for arm in arms:
    for f in ALL:
        os.environ.pop(f, None)
    for f in [x for x in arm.split(",") if x]:
        os.environ[f] = "1"
    r = PD.diff_doc(did, word, MO.measure_doc(path))
    ms = r.get("matched") if isinstance(r.get("matched"), list) else []
    bad = [m for m in ms if m.get("page_delta")]
    print(f"{arm or '(base)':32s} pass={r.get('pass')} score={r.get('score'):.4f} pcd={r.get('page_count_delta')} nbad={len(bad)}")
    for m in bad[:6]:
        print("      ", m.get("page_delta"), repr(m.get("text", m.get("word_text", "")))[:70])
