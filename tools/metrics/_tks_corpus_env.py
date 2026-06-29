import os, sys, json, glob
sys.path.insert(0,'tools/metrics')
# set env from CLI BEFORE importing (M caches RENDERER path only; env read per-subprocess)
for a in sys.argv[1:]:
    if "=" in a: k,v=a.split("=",1); os.environ[k]=v
import measure_pagination_oxi as M, pagination_diff as D
sys.stdout.reconfigure(encoding="utf-8")
DOCS='tools/golden-test/documents/docx'; WORDDIR='pipeline_data/pagination_word'; BASEDIR='pipeline_data/pagination_diff'
word_docs={os.path.splitext(os.path.basename(p))[0] for p in glob.glob(WORDDIR+'/*.json') if not os.path.basename(p).startswith('_')}
docx_map={}
for f in sorted(os.listdir(DOCS)):
    if f.lower().endswith('.docx') and not f.startswith('~$'):
        did=M.doc_id_from_filename(f)
        if did in word_docs and did not in docx_map: docx_map[did]=f
flips=[]; changed=[]; npass=0; ntot=0
for did,fn in sorted(docx_map.items()):
    try:
        oxi=M.measure_doc(os.path.join(DOCS,fn)); w=D.load_word(did)
        if w is None: continue
        r=D.diff_doc(did,w,oxi); ntot+=1; npass+= 1 if r['pass'] else 0
        bp=os.path.join(BASEDIR,did+'.json')
        if os.path.exists(bp):
            b=json.load(open(bp,encoding='utf-8'))
            if r['pass']!=b.get('pass'): flips.append((did,b.get('pass'),r['pass']))
            if abs(r['score']-b.get('score',0))>0.0003: changed.append((did,b.get('score'),r['score']))
    except Exception as e:
        print('ERR',did,str(e)[:80])
env_show={k:v for k,v in os.environ.items() if k.startswith('OXI_')}
print('ENV:',env_show)
print('n_pass=%d / %d'%(npass,ntot))
print('PASS<->FAIL flips:',len(flips))
for did,bp,np in flips:
    tag='PASS->FAIL ***' if (bp and not np) else 'FAIL->PASS'
    print('  %s %s: %s->%s'%(tag,did,bp,np))
print('score-changed (|d|>0.0003):',len(changed))
for did,bs,ns in sorted(changed,key=lambda t:(t[2]-(t[1] or 0))):
    print('  %s: %.4f -> %.4f (%+.4f)'%(did,bs or 0,ns,ns-(bs or 0)))
