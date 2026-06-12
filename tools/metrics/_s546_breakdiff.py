# -*- coding: utf-8 -*-
"""S546: diff Oxi line breaks (per-para per-line char counts) between
OXI_S546_DISABLE=1 (old model) and default (new model) for one doc.
Usage: python _s546_breakdiff.py <docx-prefix> [page]
"""
import glob, io, json, os, subprocess, sys
from collections import defaultdict

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCS = os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx')
GDI = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')


DISABLE_VAR = os.environ.get('BREAKDIFF_VAR', 'OXI_S546_DISABLE')


def dump(docx, tag, disable):
    env = dict(os.environ)
    if disable:
        env[DISABLE_VAR] = '1'
    else:
        env.pop(DISABLE_VAR, None)
    out = 'c:/tmp/_s546_bd_%s.json' % tag
    subprocess.run([GDI, docx, r'c:\tmp\_s546bd', '--dump-layout=' + out],
                   capture_output=True, env=env)
    return json.load(io.open(out, encoding='utf-8'))


def lines_of(d, page_limit):
    """(page, para_idx) -> [(y, text)] per line"""
    res = {}
    for pi, pg in enumerate(d['pages']):
        if page_limit and pi + 1 != page_limit:
            continue
        paras = defaultdict(list)
        for e in pg.get('elements', []):
            if e.get('type') == 'text':
                key = (pi + 1, e.get('para_idx'), e.get('cell_row_idx'), e.get('cell_col_idx'), e.get('cell_para_idx'))
                paras[key].append(e)
        for k, els in paras.items():
            ln = defaultdict(list)
            for e in els:
                ln[round(e['y'], 1)].append(e)
            seq = []
            for y in sorted(ln):
                row = sorted(ln[y], key=lambda e: e['x'])
                seq.append(''.join(e.get('text', '') for e in row))
            res[k] = seq
    return res


def main():
    prefix = sys.argv[1]
    page = int(sys.argv[2]) if len(sys.argv) > 2 else 0
    docx = glob.glob(os.path.join(DOCS, prefix + '*.docx'))[0]
    old = lines_of(dump(docx, 'old', True), page)
    new = lines_of(dump(docx, 'new', False), page)
    out = io.open('c:/tmp/_s546_breakdiff.txt', 'w', encoding='utf-8')
    keys = sorted(set(old) | set(new), key=lambda k: (k[0], k[1] if k[1] is not None else -1))
    ndiff = 0
    for k in keys:
        o = old.get(k, []); n = new.get(k, [])
        oc = [len(x) for x in o]; nc = [len(x) for x in n]
        if oc != nc:
            ndiff += 1
            out.write('DIFF %s old=%s new=%s\n' % (k, oc, nc))
            for i in range(max(len(o), len(n))):
                ot = o[i] if i < len(o) else ''
                nt = n[i] if i < len(n) else ''
                if ot != nt:
                    out.write('  L%d old|%s|\n  L%d new|%s|\n' % (i, ot[:46], i, nt[:46]))
    out.write('total diffs: %d (paras old=%d new=%d)\n' % (ndiff, len(old), len(new)))
    out.close()
    print('done -> c:/tmp/_s546_breakdiff.txt, diffs=%d' % ndiff)


if __name__ == '__main__':
    main()
