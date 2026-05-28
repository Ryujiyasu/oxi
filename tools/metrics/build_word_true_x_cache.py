"""S417c: batch-build the Word TRUE rendered-x reference cache for all
baseline docs, using word_true_x.measure_true_x (GetPoint-calibrated).

Replaces the Information(5) left-flow artifact x (used by
measure_pagination_word.py) with true rendered x for non-left cells, so an
x-aware Phase 2 diagnostic can validate horizontal fixes (e.g. the S412
cellMar gate). See session416/session417 memory.

Output: pipeline_data/word_true_x/<doc_id>.json  (list of paragraph recs
        with page,y,x_info5,x_true,align,text,in_table)
        pipeline_data/word_true_x/_summary.json

Per-doc isolated (own Word instance via measure_true_x); a doc that errors
is logged and skipped. Run from repo root:
    python tools/metrics/build_word_true_x_cache.py            # all baseline docs
    python tools/metrics/build_word_true_x_cache.py 1ec1 ed025 # subset by prefix
"""
from __future__ import annotations
import os, sys, json, glob, time, traceback
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(__file__))
from word_true_x import measure_true_x

REPO = r'c:\Users\ryuji\oxi-main'
DOCS_DIR = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')
SUMMARY = os.path.join(REPO, 'pipeline_data', 'pagination_diff', '_summary.json')
OUT_DIR = os.path.join(REPO, 'pipeline_data', 'word_true_x')


def baseline_doc_ids():
    with open(SUMMARY, encoding='utf-8') as f:
        s = json.load(f)
    return [d['doc_id'] for d in s['docs']]


def map_paths(doc_ids):
    out = {}
    for p in glob.glob(os.path.join(DOCS_DIR, '*.docx')):
        fn = os.path.basename(p)
        for did in doc_ids:
            if fn.startswith(did):
                out[did] = p
                break
    return out


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    prefixes = sys.argv[1:]
    ids = baseline_doc_ids()
    if prefixes:
        ids = [d for d in ids if any(d.startswith(p) for p in prefixes)]
    paths = map_paths(ids)
    summary = []
    t0 = time.time()
    for i, did in enumerate(ids):
        path = paths.get(did)
        if not path:
            print(f'[{i+1}/{len(ids)}] {did}: NO DOCX, skip', flush=True)
            summary.append({'doc_id': did, 'ok': False, 'err': 'no_docx'})
            continue
        t = time.time()
        try:
            recs = measure_true_x(path)
            outp = os.path.join(OUT_DIR, f'{did}.json')
            with open(outp, 'w', encoding='utf-8') as f:
                json.dump({'doc_id': did, 'n': len(recs), 'paragraphs': recs}, f,
                          ensure_ascii=False)
            n_true = sum(1 for r in recs if r.get('x_true') is not None)
            dt = time.time() - t
            print(f'[{i+1}/{len(ids)}] {did}: {len(recs)} paras, {n_true} x_true, {dt:.0f}s', flush=True)
            summary.append({'doc_id': did, 'ok': True, 'n': len(recs),
                            'n_x_true': n_true, 'seconds': round(dt, 1)})
        except Exception as e:
            print(f'[{i+1}/{len(ids)}] {did}: ERROR {e}', flush=True)
            traceback.print_exc()
            summary.append({'doc_id': did, 'ok': False, 'err': str(e)[:200]})
    with open(os.path.join(OUT_DIR, '_summary.json'), 'w', encoding='utf-8') as f:
        json.dump({'n_total': len(ids),
                   'n_ok': sum(1 for s in summary if s.get('ok')),
                   'total_seconds': round(time.time() - t0, 1),
                   'docs': summary}, f, ensure_ascii=False, indent=2)
    ok = sum(1 for s in summary if s.get('ok'))
    print(f'\nDONE: {ok}/{len(ids)} ok in {time.time()-t0:.0f}s -> {OUT_DIR}', flush=True)


if __name__ == '__main__':
    main()
