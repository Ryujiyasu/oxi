# -*- coding: utf-8 -*-
"""S503 confirm the S498 vc_2cell_auto cell-Y bug: a 2-col AUTO-height row, col0 short
+vAlign=center beside col1 tall (4 lines). Measure col0's content baseline Y vs Word, and
col1's lines, to confirm Oxi centers col0 too high (row-height under-count / cell-ordering).
cp932-safe: ASCII out. Matches col0 by the first CJK glyph, col1 lines by clustering."""
import os, sys, json, subprocess, tempfile, io
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
DX = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'vcenter_cellY', 'vc_2cell_auto.docx')


def word_g(dx):
    out = tempfile.mktemp(suffix='.json', dir='c:/tmp')
    subprocess.run([sys.executable, os.path.join(ROOT, 'tools', 'metrics', 'word_pdf_glyphs.py'), dx, out],
                   capture_output=True, timeout=120)
    return json.load(open(out, encoding='utf-8'))['pages'][0]['glyphs']


def oxi_g(dx, disable=False):
    e = dict(os.environ)
    if disable:
        e['OXI_S503_DISABLE'] = '1'
    gj = tempfile.mktemp(suffix='.json', dir='c:/tmp')
    subprocess.run([DW, dx, tempfile.mktemp(dir='c:/tmp'), '150', '--dump-glyphs=' + gj],
                   capture_output=True, timeout=120, env=e)
    return json.load(open(gj, encoding='utf-8'))['pages'][0]['glyphs']


def cluster_by_x(glyphs, ykey):
    # split into left-column (small x) and right-column (large x) by a gap in x
    xs = sorted(g['x'] for g in glyphs if ord(g['char'][0]) > 0x3000)
    if not xs:
        return []
    return xs


def lines_of(glyphs, ykey, xmin=None, xmax=None):
    gs = [g for g in glyphs if ord(g['char'][0]) > 0x3000]
    if xmin is not None:
        gs = [g for g in gs if g['x'] >= xmin]
    if xmax is not None:
        gs = [g for g in gs if g['x'] < xmax]
    rows = {}
    for g in gs:
        k = round(g[ykey], 0)
        rows.setdefault(k, []).append(g)
    out = []
    for k in sorted(rows):
        r = sorted(rows[k], key=lambda g: g['x'])
        out.append((k, r[0]['x'], ''.join(g['char'] for g in r)))
    return out


def main():
    wg = word_g(DX)
    og = oxi_g(DX)
    # find the x gap between col0 and col1
    L = ['S503 vc_2cell_auto: Word vs Oxi cell lines (y=baseline)']
    # heuristic column split: median x
    allx = sorted(g['x'] for g in wg if ord(g['char'][0]) > 0x3000)
    split = (min(allx) + max(allx)) / 2 if allx else 200
    L.append('col split x ~ %.0f' % split)
    for tag, g, yk in [('WORD', wg, 'y'), ('OXI ', og, 'baseline')]:
        L.append('\n%s col0 (x<%.0f):' % (tag, split))
        for k, x, t in lines_of(g, yk, xmax=split):
            L.append('   y=%6.1f x=%6.1f  %s' % (k, x, t[:16]))
        L.append('%s col1 (x>=%.0f):' % (tag, split))
        for k, x, t in lines_of(g, yk, xmin=split):
            L.append('   y=%6.1f x=%6.1f  %s' % (k, x, t[:16]))
    with io.open('c:/tmp/_s503_vc2_out.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print('wrote c:/tmp/_s503_vc2_out.txt')


if __name__ == '__main__':
    main()
