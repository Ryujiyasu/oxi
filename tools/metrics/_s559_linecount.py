# -*- coding: utf-8 -*-
"""S559 — render 3a4f under OXI_S559_CELLMAR = (unset / tcwgt / all) and diff
per-paragraph LINE COUNTS (distinct y per cell-para group). This separates the
GENUINE wrap changes (cellMar reservation actually changing a wrap) from the
PAGINATION CASCADE ({1:1323} = downstream page shifts, not 1323 real changes).

Decisive question: under tcwgt, does ONLY ⑦ (and a handful) change line count,
or do many cells spuriously over-wrap? If few + ⑦ included → narrow gate is safe.
"""
import json
import os
import subprocess
import sys
import tempfile
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')
REPO = r'c:\Users\ryuji\oxi-main'
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')
DOCX = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', '3a4f9fbe1a83_001620506.docx')
ANCHOR = u'常に整理整頓'


def render(mode):
    env = dict(os.environ)
    if mode is None:
        env.pop('OXI_S559_CELLMAR', None)
    else:
        env['OXI_S559_CELLMAR'] = mode
    with tempfile.TemporaryDirectory(prefix='s559_') as tmp:
        out_prefix = os.path.join(tmp, 'p_')
        dump = os.path.join(tmp, 'layout.json')
        r = subprocess.run([RENDERER, DOCX, out_prefix, '--dump-layout=' + dump],
                           capture_output=True, text=True, timeout=300, env=env)
        if not os.path.exists(dump):
            print('RENDER FAILED mode=%s rc=%d' % (mode, r.returncode))
            print(r.stderr[-2000:])
            sys.exit(1)
        with open(dump, encoding='utf-8') as f:
            return json.load(f)


def linecounts(dump):
    """(para_idx, cpi, cri, cci) -> set of y values (lines); + text prefix."""
    lines = defaultdict(set)
    texts = {}
    npages = {}
    for page in dump.get('pages', []):
        pg = page['page']
        for el in page.get('elements', []):
            if el.get('type') != 'text':
                continue
            key = (el.get('para_idx'), el.get('cell_para_idx'),
                   el.get('cell_row_idx'), el.get('cell_col_idx'))
            lines[key].add(round(el['y'], 1))
            texts.setdefault(key, '')
            if len(texts[key]) < 24:
                texts[key] += el.get('text', '')
            npages[key] = pg
    return {k: (len(v), texts.get(k, ''), npages.get(k)) for k, v in lines.items()}


def main():
    modes = sys.argv[1:] or ['tcwgt', 'just', 'justl0']
    print('rendering OFF...')
    off = linecounts(render(None))
    results = []
    for mode in modes:
        print('rendering %s...' % mode)
        results.append((mode, linecounts(render(mode))))

    for label, m in results:
        changed = []
        for k, (n, txt, pg) in off.items():
            if k in m and m[k][0] != n:
                changed.append((k, n, m[k][0], txt, pg))
        # also keys new to m (shouldn't happen) — ignore
        print('\n===== %s vs OFF: %d paras changed line count =====' % (label, len(changed)))
        anchor_hit = False
        for k, n0, n1, txt, pg in sorted(changed, key=lambda x: (x[1] - x[2])):
            mark = ''
            if ANCHOR in txt:
                mark = '  <== ⑦'
                anchor_hit = True
            print('   %d -> %d  (p%s)  %r%s' % (n0, n1, pg, txt, mark))
        print('   ⑦ changed: %s' % anchor_hit)
        # net extra lines
        net = sum(n1 - n0 for _, n0, n1, _, _ in changed)
        print('   net line delta: %+d' % net)


if __name__ == '__main__':
    main()
