# -*- coding: utf-8 -*-
"""S519: per-CELL first-line baseline alignment within a table row, Word PDF vs Oxi dwrite dump.
For a tokumei doc, find table-row bands, cluster glyphs into cells by x-gaps, and report each
cell's first baseline (Word, Oxi, Oxi-Word). If the per-cell deltas are UNIFORM across a row =
a uniform table shift (d4d126 class, S498); if they VARY across cells = a real per-cell
within-row misalignment (the NEW S518-audit angle). cp932-safe: UTF-8 file, results to file."""
import os, sys, json, subprocess, io
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')

def oxi_glyphs(docx):
    pre = os.path.join('c:/tmp', 's519_' + os.path.splitext(os.path.basename(docx))[0][:14])
    gj = pre + '_g.json'
    subprocess.run([EXE, os.path.abspath(docx), pre, '72', '--dump-glyphs=' + gj], capture_output=True, text=True)
    return json.load(open(gj, encoding='utf-8'))['pages']

def word_glyphs(docx):
    import win32com.client, pythoncom, fitz
    pdf = os.path.join('c:/tmp', os.path.basename(docx)[:14] + '_w.pdf')
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(os.path.abspath(docx), ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    pages = []
    for pg in fitz.open(pdf):
        gs = []
        for blk in pg.get_text('rawdict').get('blocks', []):
            for ln in blk.get('lines', []):
                for sp in ln.get('spans', []):
                    for c in sp.get('chars', []):
                        if c['c'].strip():
                            gs.append({'char': c['c'], 'x': c['origin'][0], 'baseline': c['origin'][1]})
        pages.append(gs)
    return pages

def cluster_cells(band):
    """cluster glyphs (sorted by x) into cells by x-gap > 14pt."""
    band = sorted(band, key=lambda g: g['x'])
    cells = []; cur = [band[0]]
    for g in band[1:]:
        if g['x'] - cur[-1]['x'] > 14:
            cells.append(cur); cur = [g]
        else:
            cur.append(g)
    cells.append(cur)
    return cells

def first_baselines_in_row(glyphs, y_lo, y_hi):
    band = [g for g in glyphs if y_lo <= g['baseline'] <= y_hi]
    if not band:
        return []
    cells = cluster_cells(band)
    return [(min(c, key=lambda g: g['x'])['x'], c[0]['baseline'],
             ''.join(g['char'] for g in sorted(c, key=lambda g: g['x']))[:6]) for c in cells]

def main():
    docx = sys.argv[1] if len(sys.argv) > 1 else \
        [os.path.join(ROOT, 'tools/golden-test/documents/docx', f) for f in
         os.listdir(os.path.join(ROOT, 'tools/golden-test/documents/docx')) if f.startswith('d4d126')][0]
    page = int(sys.argv[2]) if len(sys.argv) > 2 else 4
    ox = oxi_glyphs(docx); wd = word_glyphs(docx)
    op = ox[page - 1]['glyphs'] if page - 1 < len(ox) else []
    wp = wd[page - 1] if page - 1 < len(wd) else []
    L = ['S519 per-cell row baseline: %s p%d' % (os.path.basename(docx), page)]
    # scan rows: bucket Word glyphs by baseline (round 0.5), pick rows with >=2 cells
    import collections
    wb = collections.defaultdict(list)
    for g in wp:
        wb[round(g['baseline'] * 2) / 2].append(g)
    rows = sorted(wb)
    for wbl in rows[:30]:
        wcells = first_baselines_in_row(wp, wbl - 1.0, wbl + 1.0)
        if len(wcells) < 2:
            continue
        # matching Oxi row: nearest oxi baseline band
        ocells = first_baselines_in_row(op, wbl - 4.0, wbl + 4.0)
        # align cells by x
        L.append('--- Word row bl~%.1f : %d cells' % (wbl, len(wcells)))
        for (wx, wbl2, wt) in wcells:
            # nearest oxi cell by x
            best = min(ocells, key=lambda o: abs(o[0] - wx)) if ocells else None
            if best and abs(best[0] - wx) < 12:
                L.append('   x~%5.1f W_bl=%6.2f O_bl=%6.2f  O-W=%+5.2f  %r' % (wx, wbl2, best[1], best[1] - wbl2, wt))
            else:
                L.append('   x~%5.1f W_bl=%6.2f O_bl=  ?     (no oxi match) %r' % (wx, wbl2, wt))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s519_out.txt', 'w', encoding='utf-8').write(txt + '\n')
    print('wrote c:/tmp/_s519_out.txt ; rows reported:', txt.count('--- Word row'))

if __name__ == '__main__':
    main()
