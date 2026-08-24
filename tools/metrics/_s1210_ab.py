# -*- coding: utf-8 -*-
"""Phase-1 A/B where BOTH arms carry a base env set (the derived-cell bundle).

`_golden_ab_envs.py` toggles every flag it is given together, so it can only
compare "all off" against "all on". Promoting one flag INSIDE the bundle needs
the bundle held fixed in both arms.

    python _s1210_ab.py OXI_S1210 [doc_id ...]
"""
import json, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import measure_pagination_oxi as MO
import pagination_diff as PD

REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
WORD_DIR = os.path.join(REPO, "pipeline_data", "pagination_word")
DOCS_DIR = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
BASE = {"OXI_CELLLAW": "1", "OXI_YAKUCOMP": "1", "OXI_AUTOSPACE2": "1",
        "OXI_S1201": "1", "OXI_S1155": "1"}
FLAGS = [f for f in sys.argv[1].split(",") if f]
LOG = "C:/tmp/%s_inbundle.log" % "_".join(FLAGS)

by_id = {}
for f in sorted(os.listdir(DOCS_DIR)):
    if f.endswith(".docx") and not f.startswith("~$"):
        by_id[f[:-5]] = os.path.join(DOCS_DIR, f)

ids = sys.argv[2:] or sorted(d[:-5] for d in os.listdir(WORD_DIR)
                             if d.endswith(".json") and not d.startswith("_") and len(d) > 5)


def arm(on):
    os.environ.update(BASE)
    for flag in FLAGS:
        name, _, value = flag.partition("=")
        os.environ.pop(name, None)
        if on:
            os.environ[name] = value or "1"


n_a = n_b = 0
with open(LOG, "a", encoding="utf-8") as log:
    for did in ids:
        path = by_id.get(did) or next((v for k, v in by_id.items()
                                       if k.split("_")[0] == did), None)
        if path is None:
            continue
        word = json.load(open(os.path.join(WORD_DIR, did + ".json"), encoding="utf-8"))
        out = {}
        try:
            for on in (False, True):
                arm(on)
                out[on] = PD.diff_doc(did, word, MO.measure_doc(path))
        except Exception as exc:  # noqa: BLE001
            log.write("%s\tERR\t%s\n" % (did, exc)); log.flush(); continue
        a, b = out[False], out[True]
        n_a += bool(a["pass"]); n_b += bool(b["pass"])
        flip = ("  *** PASS->FAIL ***" if a["pass"] and not b["pass"]
                else "  *** FAIL->PASS ***" if b["pass"] and not a["pass"] else "")
        line = "%s\t%s\t%.4f\t%d\t%s\t%.4f\t%d%s" % (
            did, a["pass"], a["score"], a["page_count_delta"],
            b["pass"], b["score"], b["page_count_delta"], flip)
        log.write(line + "\n"); log.flush()
        if flip or abs(a["score"] - b["score"]) > 1e-9:
            print(line, flush=True)
print("PASS %d -> %d (n=%d)" % (n_a, n_b, len(ids)))
