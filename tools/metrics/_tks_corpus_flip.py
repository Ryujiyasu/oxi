import os, sys, json, glob
sys.path.insert(0,'tools/metrics')
import measure_pagination_oxi as M, pagination_diff as D
DOCS='tools/golden-test/documents/docx'; WORDDIR='pipeline_data/pagination_word'; BASEDIR='pipeline_data/pagination_diff'
word_docs={os.path.splitext(os.path.basename(p))[0] for p in glob.glob(WORDDIR+'/*.json')}
docx_map={}
for f in sorted(os.listdir(DOCS)):
    if f.lower().endswith('.docx') and not f.startswith('~$'):
        did=M.doc_id_from_filename(f)
        if did in word_docs and did not in docx_map: docx_map[did]=f
os.environ['OXI_S590']='1'
flips=[]; changed=[]
for did,fn in sorted(docx_map.items()):
    try:
        oxi=M.measure_doc(os.path.join(DOCS,fn)); w=D.load_word(did)
        if w is None: continue
        r=D.diff_doc(did,w,oxi)
        bp=os.path.join(BASEDIR,did+'.json')
        if os.path.exists(bp):
            b=json.load(open(bp,encoding='utf-8'))
            if r['pass']!=b.get('pass'): flips.append((did,b.get('pass'),r['pass']))
            if abs(r['score']-b.get('score',0))>0.0003: changed.append((did,b.get('score'),r['score']))
    except Exception as e:
        print('ERR',did,str(e)[:60])
print('PASS<->FAIL flips:',len(flips))
for did,bp,np in flips:
    tag='PASS->FAIL ***' if (bp and not np) else 'FAIL->PASS'
    print('  %s %s: %s->%s'%(tag,did,bp,np))
print('score-changed (|d|>0.0003):',len(changed))
for did,bs,ns in changed:
    print('  %s: %s -> %.4f (%+.4f)'%(did,bs,ns,ns-(bs or 0)))
