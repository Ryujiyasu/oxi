# -*- coding: utf-8 -*-
"""S518 corpus-wide within-line baseline-split audit. For every baseline docx, find visual lines
whose glyphs span >1 baseline (>1.5pt). Classify each split:
  SAME-SIZE  = all sub-baselines share the line's dominant font_size -> SUSPICIOUS (a baseline bug
               like the S517 list-marker; Word would share the baseline)
  RAISED-SM  = the higher sub-baseline group is a SMALLER font -> likely legit superscript/ref
Report docs with SAME-SIZE splits first (the real-bug candidates). cp932-safe."""
import os, sys, json, subprocess, io, collections, glob
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
DOCX_DIR = os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx')

def scan(docx):
    pre = os.path.join('c:/tmp', 's518_' + os.path.splitext(os.path.basename(docx))[0][:20])
    gj = pre + '_g.json'
    try:
        subprocess.run([EXE, os.path.abspath(docx), pre, '72', '--dump-glyphs=' + gj],
                       capture_output=True, text=True, timeout=120)
    except Exception:
        return None
    if not os.path.exists(gj):
        return None
    data = json.load(open(gj, encoding='utf-8'))
    same_splits = []   # (page, span, mainfs, txt)
    raised_sm = 0
    for pi, pg in enumerate(data['pages']):
        g = [x for x in pg['glyphs'] if x['char'].strip()]
        g.sort(key=lambda c: c['baseline'])
        bands = []; cur = []
        for x in g:
            if cur and (x['baseline'] - min(c['baseline'] for c in cur)) > 6:
                bands.append(cur); cur = []
            cur.append(x)
        if cur: bands.append(cur)
        for band in bands:
            bys = collections.defaultdict(list)
            for c in band:
                bys[round(c['baseline'] * 2) / 2].append(c)
            distinct = sorted(bys)
            if len(distinct) < 2 or (distinct[-1] - distinct[0]) <= 1.5:
                continue
            # dominant font size = mode over the whole band
            dom_fs = collections.Counter(round(c['font_size'], 1) for c in band).most_common(1)[0][0]
            # the higher (smaller-y) group:
            hi = bys[distinct[0]]
            hi_fs = collections.Counter(round(c['font_size'], 1) for c in hi).most_common(1)[0][0]
            # the lower (main) group:
            lo = bys[distinct[-1]]
            lo_fs = collections.Counter(round(c['font_size'], 1) for c in lo).most_common(1)[0][0]
            if hi_fs < lo_fs - 0.5:
                raised_sm += 1   # smaller glyph raised = likely superscript
            else:
                txt = ''.join(c['char'] for c in sorted(band, key=lambda c: c['x']))[:20]
                same_splits.append((pi + 1, round(distinct[-1] - distinct[0], 1), dom_fs, txt))
    return same_splits, raised_sm

def main():
    docs = sorted(glob.glob(os.path.join(DOCX_DIR, '*.docx')))
    if len(sys.argv) > 1:
        docs = docs[:int(sys.argv[1])]
    rows = []
    for i, dx in enumerate(docs):
        r = scan(dx)
        if r is None:
            continue
        same, raised = r
        if same:
            rows.append((os.path.basename(dx)[:28], len(same), raised, same[:4]))
    rows.sort(key=lambda x: -x[1])
    L = ['S518 corpus baseline-split audit (SAME-SIZE splits = suspicious bugs)']
    L.append('docs scanned=%d, docs with SAME-SIZE splits=%d' % (len(docs), len(rows)))
    L.append('%-30s same raised  examples(page,span,fs,text)' % 'doc')
    for name, ns, raised, ex in rows:
        exs = ' '.join('p%d/%.1fpt/%sfs/%r' % (e[0], e[1], e[2], e[3]) for e in ex)
        L.append('%-30s %4d %5d  %s' % (name, ns, raised, exs))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s518_audit.txt', 'w', encoding='utf-8').write(txt + '\n')
    print('docs with SAME-SIZE splits: %d / %d  -> c:/tmp/_s518_audit.txt' % (len(rows), len(docs)))

if __name__ == '__main__':
    main()
