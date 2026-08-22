# -*- coding: utf-8 -*-
"""Two-arm Phase-1 census over the golden set for a COMBINATION of flags.

`_s1189_golden_ab.py` takes one flag; a bundle whose members only work together
needs all of them in the B arm (S1192 needs S1195 to keep ed025 passing).

  python _golden_ab_envs.py OXI_S1192,OXI_S1195 [start] [count]

Appends one line per doc to C:/tmp/<joined>_golden.log so a kill keeps progress.
A flag named *_DISABLE is inverted (set in the A arm), matching
pipeline_data/en_benchmark/_ab_envs.py.
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

FLAGS = [f for f in sys.argv[1].split(",") if f]
START = int(sys.argv[2]) if len(sys.argv) > 2 else 0
COUNT = int(sys.argv[3]) if len(sys.argv) > 3 else 10 ** 9
LOG = "C:/tmp/%s_golden.log" % "_".join(FLAGS)


def arm(on):
    for flag in FLAGS:
        name, _, value = flag.partition("=")
        invert = name.endswith("_DISABLE")
        os.environ.pop(name, None)
        if on != invert:
            os.environ[name] = value or "1"


by_id = {}
for f in sorted(os.listdir(DOCS_DIR)):
    if not f.endswith(".docx") or f.startswith("~$"):
        continue
    by_id[f[:-5]] = os.path.join(DOCS_DIR, f)

ids = sorted(d[:-5] for d in os.listdir(WORD_DIR) if d.endswith(".json")
             and not d.startswith("_"))
todo = ids[START:START + COUNT]
n_a = n_b = 0
with open(LOG, "a", encoding="utf-8") as log:
    for did in todo:
        path = by_id.get(did)
        if path is None:
            cands = [v for k, v in by_id.items() if k.startswith(did.split("_")[0])]
            path = cands[0] if cands else None
        if path is None:
            log.write("%s\tNO_DOCX\n" % did); log.flush(); continue
        word = json.load(open(os.path.join(WORD_DIR, did + ".json"), encoding="utf-8"))
        out = {}
        try:
            for on in (False, True):
                arm(on)
                out[on] = PD.diff_doc(did, word, MO.measure_doc(path))
        except Exception as exc:  # noqa: BLE001
            log.write("%s\tERR\t%s\n" % (did, exc)); log.flush(); continue
        finally:
            for flag in FLAGS:
                os.environ.pop(flag.partition("=")[0], None)
        a, b = out[False], out[True]
        n_a += bool(a["pass"]); n_b += bool(b["pass"])
        flip = ""
        if a["pass"] and not b["pass"]:
            flip = "  *** PASS->FAIL ***"
        elif b["pass"] and not a["pass"]:
            flip = "  *** FAIL->PASS ***"
        line = ("%s\t%s\t%.4f\t%d\t%s\t%.4f\t%d%s"
                % (did, a["pass"], a["score"], a["page_count_delta"],
                   b["pass"], b["score"], b["page_count_delta"], flip))
        log.write(line + "\n"); log.flush()
        if flip or abs(a["score"] - b["score"]) > 1e-9:
            print(line, flush=True)
print("PASS %d -> %d (n=%d)" % (n_a, n_b, len(todo)))
