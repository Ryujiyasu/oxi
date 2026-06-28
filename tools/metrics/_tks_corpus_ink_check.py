# -*- coding: utf-8 -*-
import os, sys, json, glob
sys.path.insert(0,'tools/metrics')
import measure_pagination_oxi as M, pagination_diff as D
DOCS='tools/golden-test/documents/docx'
WORDDIR='pipeline_data/pagination_word'
BASEDIR='pipeline_data/pagination_diff'
# docs with a word baseline = the Phase-1 corpus
word_docs={os.path.splitext(os.path.basename(p))[0] for p in glob.glob(WORDDIR+'/*.json')}
# map doc_id -> docx
docx_map={}
for f in sorted(os.listdir(DOCS)):
    if f.lower().endswith('.docx') and not f.startswith('~$'):
        did=M.doc_id_from_filename(f)
        if did in word_docs and did not in docx_map: docx_map[did]=f
print(f"corpus docs with word baseline: {len(docx_map)}")
os.environ['OXI_TGINK_K']='1.0'
flips=[]; changed=[]; errors=[]
for i,(did,fn) in enumerate(sorted(docx_map.items())):
    try:
        oxi=M.measure_doc(os.path.join(DOCS,fn))
        w=D.load_word(did)
        if w is None: continue
        r=D.diff_doc(did,w,oxi)
        # baseline
        bp=os.path.join(BASEDIR,did+'.json')
        if os.path.exists(bp):
            b=json.load(open(bp,encoding='utf-8'))
            bpass=b.get('pass'); bscore=b.get('score')
        else:
            bpass=None; bscore=None
        if bpass is not None and r['pass']!=bpass:
            flips.append((did,bpass,r['pass'],bscore,r['score']))
        if bscore is not None and abs(r['score']-bscore)>0.0005:
            changed.append((did,bscore,r['score']))
    except Exception as e:
        errors.append((did,str(e)[:80]))
    if (i+1)%20==0: print(f"  ...{i+1}/{len(docx_map)}")
print(f"\n=== RESULT (OXI_TGINK_K=1.0 vs committed baseline) ===")
print(f"PASS<->FAIL flips: {len(flips)}")
for did,bp,np,bs,ns in flips:
    tag='PASS->FAIL ***' if (bp and not np) else 'FAIL->PASS'
    print(f"  {tag} {did}: {bs}->{ns}")
print(f"\nscore-changed docs (|d|>0.0005): {len(changed)}")
for did,bs,ns in sorted(changed,key=lambda x:x[2]-x[1]):
    print(f"  {did}: {bs} -> {ns:.4f} ({ns-bs:+.4f})")
if errors: print("errors:",errors)
