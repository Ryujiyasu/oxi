# -*- coding: utf-8 -*-
"""Two-arm Phase-1 census over the golden set (pipeline_data/pagination_word).

  python _s1189_golden_ab.py OXI_S1189 [start] [count]

Appends one line per doc to C:/tmp/<FLAG>_golden.log so a kill keeps progress.
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

FLAG = sys.argv[1]
START = int(sys.argv[2]) if len(sys.argv) > 2 else 0
COUNT = int(sys.argv[3]) if len(sys.argv) > 3 else 10**9
LOG = f"C:/tmp/{FLAG}_golden.log"

by_id = {}
for f in sorted(os.listdir(DOCS_DIR)):
    if not f.endswith(".docx") or f.startswith("~$"):
        continue
    by_id[f[:-5]] = os.path.join(DOCS_DIR, f)

ids = sorted(d[:-5] for d in os.listdir(WORD_DIR) if d.endswith(".json"))
todo = ids[START:START + COUNT]
with open(LOG, "a", encoding="utf-8") as log:
    for did in todo:
        path = by_id.get(did)
        if path is None:
            cands = [v for k, v in by_id.items() if k.startswith(did.split("_")[0])]
            path = cands[0] if cands else None
        if path is None:
            log.write(f"{did}\tNO_DOCX\n"); log.flush(); continue
        word = json.load(open(os.path.join(WORD_DIR, did + ".json"), encoding="utf-8"))
        out = {}
        try:
            for arm in ("A", "B"):
                os.environ.pop(FLAG, None)
                if arm == "B":
                    os.environ[FLAG] = "1"
                out[arm] = PD.diff_doc(did, word, MO.measure_doc(path))
        except Exception as exc:
            log.write(f"{did}\tERR\t{exc}\n"); log.flush(); continue
        finally:
            os.environ.pop(FLAG, None)
        a, b = out["A"], out["B"]
        line = (f"{did}\t{a.get('pass')}\t{a.get('score'):.4f}\t{a.get('page_count_delta')}"
                f"\t{b.get('pass')}\t{b.get('score'):.4f}\t{b.get('page_count_delta')}")
        log.write(line + "\n"); log.flush()
        print(line)
