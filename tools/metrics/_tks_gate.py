# -*- coding: utf-8 -*-
# Gate-accurate per-doc official pagination score (env-respecting).
# Usage: python tools/metrics/_tks_gate.py [doc_id] [KEY=VAL ...]
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import measure_pagination_oxi as M
import pagination_diff as D
sys.stdout.reconfigure(encoding="utf-8")

doc_id = "tokyoshugyo"
for a in sys.argv[1:]:
    if "=" in a:
        k,v=a.split("=",1); os.environ[k]=v
    elif not a.startswith("--"):
        doc_id=a

# find docx
fname=None
for f in os.listdir(M.DOCS_DIR):
    if f.lower().endswith(".docx") and M.doc_id_from_filename(f).startswith(doc_id):
        fname=f; break
if not fname:
    print("no docx for", doc_id); sys.exit(1)
oxi = M.measure_doc(os.path.join(M.DOCS_DIR, fname))
oxi["doc_id"]=M.doc_id_from_filename(fname)
word = D.load_word(M.doc_id_from_filename(fname))
res = D.diff_doc(M.doc_id_from_filename(fname), word, oxi)
env_show = {k:v for k,v in os.environ.items() if k.startswith("OXI_")}
print(f"doc={res['doc_id']} env={env_show}")
print(f"  pass={res['pass']} score={res['score']} pcd={res['page_count_delta']} "
      f"oxi_pages={res['oxi_n_pages']} word_pages={res['word_n_pages']}")
print(f"  n_matched={res['n_matched']} n_zero={res['n_page_zero']} "
      f"n_neg={res['n_page_negative']} n_pos={res['n_page_positive']} unmatched={res['n_unmatched']}")
print(f"  delta_hist={res['delta_histogram']}")
if "--rows" in sys.argv:
    for m in res["matches"]:
        if m["matched"] and m["page_delta"]!=0:
            print(f"   wi={m['word_i']:>4} wp={m['word_page']:>2} d={m['page_delta']:+d}  {m['text'][:24]}")
