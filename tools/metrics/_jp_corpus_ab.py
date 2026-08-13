# -*- coding: utf-8 -*-
"""A/B a layout env flag set across the JP Phase-1 corpus, in memory.

The EN twin of this lives at pipeline_data/en_benchmark/_ab_envs.py; this one
walks the golden-test corpus (every doc that has cached Word truth under
pipeline_data/pagination_word/) so a Latin-scoped rule can be shown NOT to move
the JP sentinel without overwriting the pagination_oxi cache.

  python _jp_corpus_ab.py OXI_S1112_DISABLE,OXI_S1091_DISABLE,OXI_S1074_DISABLE
    -> A arm sets the *_DISABLE flags (the old default), B arm is the shipped one.
"""
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


def docx_for(doc_id: str):
    for f in sorted(os.listdir(DOCS_DIR)):
        if f.endswith(".docx") and MO.doc_id_from_filename(f) == doc_id:
            return os.path.join(DOCS_DIR, f)
    return None


def measure(path, flags, on):
    for flag in flags:
        name, _, value = flag.partition("=")
        invert = name.endswith("_DISABLE")
        os.environ.pop(name, None)
        if on != invert:
            os.environ[name] = value or "1"
    try:
        return MO.measure_doc(path)
    except Exception as exc:  # noqa: BLE001
        return {"error": str(exc)}
    finally:
        for flag in flags:
            os.environ.pop(flag.partition("=")[0], None)


def main() -> None:
    flags = sys.argv[1].split(",")
    ids = sorted(os.path.splitext(f)[0] for f in os.listdir(WORD_DIR)
                 if f.endswith(".json") and not f.startswith("_"))
    na = nb = n = 0
    for did in ids:
        path = docx_for(did)
        if not path:
            continue
        word = json.load(open(os.path.join(WORD_DIR, did + ".json"), encoding="utf-8"))
        a, b = measure(path, flags, False), measure(path, flags, True)
        if "error" in a or "error" in b:
            print(f"  ERROR {did}", flush=True)
            continue
        da, db = PD.diff_doc(did, word, a), PD.diff_doc(did, word, b)
        n += 1
        na += bool(da["pass"])
        nb += bool(db["pass"])
        if abs(da["score"] - db["score"]) > 1e-9 or \
                da["page_count_delta"] != db["page_count_delta"]:
            flip = ""
            if da["pass"] and not db["pass"]:
                flip = "  *** PASS->FAIL ***"
            elif db["pass"] and not da["pass"]:
                flip = "  *** FAIL->PASS ***"
            print(f"  {did:24s} {da['score']:.4f} -> {db['score']:.4f}"
                  f"  pcd {da['page_count_delta']:+d} -> {db['page_count_delta']:+d}{flip}",
                  flush=True)
    print(f"\nJP corpus: PASS {na} -> {nb}  (n={n})")


if __name__ == "__main__":
    main()
