"""Chinese dev pool: deterministic selection + Word truth PDFs (2026-11-09, user decision A:
documents whose faces this machine lacks stay in -- Word's own substitution is the truth).

Selection (no quality look, same spirit as FROZEN_SELECTION_RULE.md): from
pipeline_data/docx_corpus/zh/_census.json take the documents of one script (hans / hant),
per type the first 5 in SHA-256 ascending order; if a type has fewer, top up from the
remaining documents of that script in SHA-256 order across types until N.
Truth: Word COM -> ExportAsFixedFormat PDF, one document at a time (never concurrent COM).

Usage: python zh_dev_truth.py <hans|hant> [N=50]
Writes pipeline_data/zh_benchmark/dev_<script><N>/{ids.json, word_pdf/<type>__<sha16>.pdf, word.log}
"""
import sys, os, json, time

sys.stdout.reconfigure(encoding='utf-8')
REPO = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
script = sys.argv[1]
N = int(sys.argv[2]) if len(sys.argv) > 2 else 50
corpus = os.path.join(REPO, 'pipeline_data', 'docx_corpus', 'zh')
rows = [r for r in json.load(open(os.path.join(corpus, '_census.json'), encoding='utf8'))
        if r.get('kind') == script]
by_type = {}
for r in sorted(rows, key=lambda r: r['doc']):
    by_type.setdefault(r['doc'].split('/')[0], []).append(r['doc'])
pick = []
for t in sorted(by_type):
    pick += by_type[t][:5]
rest = sorted((d for t in by_type for d in by_type[t][5:]), key=lambda d: d.split('/')[1])
pick += rest[:max(0, N - len(pick))]
pick = pick[:N]

out = os.path.join(REPO, 'pipeline_data', 'zh_benchmark', f'dev_{script}{N}')
os.makedirs(os.path.join(out, 'word_pdf'), exist_ok=True)
json.dump(pick, open(os.path.join(out, 'ids.json'), 'w', encoding='utf8'), indent=1)
print(script, 'picked', len(pick))

import win32com.client
w = win32com.client.DispatchEx('Word.Application')
w.Visible = False
w.DisplayAlerts = 0
log = open(os.path.join(out, 'word.log'), 'a', encoding='utf8')
try:
    for i, d in enumerate(pick, 1):
        t, f = d.split('/')
        stem = f'{t}__{f[:-5]}'
        dst = os.path.join(out, 'word_pdf', stem + '.pdf')
        if os.path.exists(dst):
            continue
        t0 = time.time()
        try:
            doc = w.Documents.Open(os.path.join(corpus, t, f), ReadOnly=True, AddToRecentFiles=False)
            doc.ExportAsFixedFormat(dst, 17)
            doc.Close(False)
            msg = f'[WORD {i:2d}/{len(pick)}] {stem} OK {time.time() - t0:.1f}s'
        except Exception as e:
            msg = f'[WORD {i:2d}/{len(pick)}] {stem} FAIL {e}'
        print(msg); log.write(msg + '\n'); log.flush()
finally:
    w.Quit()
