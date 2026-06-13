# -*- coding: utf-8 -*-
"""S559 — structural analysis of 3a4f's tables (NO COM, NO build).

Goal: find the table-STRUCTURAL discriminator that makes Word reserve the
default cellMar (108tw/side) for the ⑦ 遵守事項 cell (para 2234) but NOT for
the ~1323 other single-cell body paras (which regressed {1:1323} when a
universal cellMar subtraction was tried, S559).

Walks every w:tbl (incl. nested), records tblStyle/tblW/tblLayout/tblInd/
tblCellMar/nrows + per-row ncells + per-cell tcW + gridCol widths. Locates the
table+row holding the ⑦ anchor, dumps it in full, then buckets ALL single-cell
ROWS by structural signature so the ⑦ bucket can be told apart from the 1323.
"""
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict

DOCX = r'c:\tmp\3a4f9f.docx'
ANCHOR = u'常に整理整頓'  # 常に整理整頓
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
sys.stdout.reconfigure(encoding='utf-8')


def q(tag):
    return W + tag


def attr(el, name):
    return el.get(W + name) if el is not None else None


def tbl_text(tbl, limit=24):
    return ''.join(t.text or '' for t in tbl.iter(q('t')))[:limit]


def cell_text(tc, limit=40):
    return ''.join(t.text or '' for t in tc.iter(q('t')))[:limit]


def grid_cols(tbl):
    g = tbl.find(q('tblGrid'))
    if g is None:
        return []
    return [attr(c, 'w') for c in g.findall(q('gridCol'))]


def tbl_signature(tbl):
    pr = tbl.find(q('tblPr'))
    style = attr(pr.find(q('tblStyle')), 'val') if pr is not None and pr.find(q('tblStyle')) is not None else None
    tblw_el = pr.find(q('tblW')) if pr is not None else None
    tblw = (attr(tblw_el, 'type'), attr(tblw_el, 'w')) if tblw_el is not None else (None, None)
    layout_el = pr.find(q('tblLayout')) if pr is not None else None
    layout = attr(layout_el, 'type') if layout_el is not None else None
    ind_el = pr.find(q('tblInd')) if pr is not None else None
    tblind = attr(ind_el, 'w') if ind_el is not None else None
    cm = pr.find(q('tblCellMar')) if pr is not None else None
    has_cellmar = cm is not None
    cm_vals = None
    if has_cellmar:
        cm_vals = {}
        for side in ('top', 'left', 'bottom', 'right'):
            s = cm.find(q(side))
            if s is not None:
                cm_vals[side] = (attr(s, 'type'), attr(s, 'w'))
    rows = tbl.findall(q('tr'))
    nrows = len(rows)
    cols = grid_cols(tbl)
    return dict(style=style, tblw=tblw, layout=layout, tblind=tblind,
                has_cellmar=has_cellmar, cm_vals=cm_vals, nrows=nrows,
                gridcols=cols, ncols=len(cols))


def row_cells(tr):
    return tr.findall(q('tc'))


def cell_tcw(tc):
    pr = tc.find(q('tcPr'))
    if pr is None:
        return (None, None)
    w = pr.find(q('tcW'))
    if w is None:
        return (None, None)
    return (attr(w, 'type'), attr(w, 'w'))


def main():
    with zipfile.ZipFile(DOCX) as z:
        xml = z.read('word/document.xml')
    root = ET.fromstring(xml)
    body = root.find(q('body'))

    # collect all tables (top-level + nested) by iter
    all_tbls = list(root.iter(q('tbl')))
    print('TOTAL tables in doc:', len(all_tbls))

    # locate the ⑦ table + row
    target_tbl = None
    target_row = None
    for tbl in all_tbls:
        for tr in tbl.findall(q('tr')):
            txt = ''.join(t.text or '' for t in tr.iter(q('t')))
            if ANCHOR in txt:
                target_tbl = tbl
                target_row = tr
                break
        if target_tbl is not None:
            break

    print('\n===== ⑦ TABLE (contains anchor) =====')
    if target_tbl is None:
        print('  ANCHOR NOT FOUND in any table!')
    else:
        sig = tbl_signature(target_tbl)
        print('  sig:', sig)
        print('  row preview (first 8 rows):')
        for i, tr in enumerate(target_tbl.findall(q('tr'))[:8]):
            cells = row_cells(tr)
            tcws = [cell_tcw(c) for c in cells]
            print('    r%d  ncells=%d  tcW=%s  text=%r'
                  % (i, len(cells), tcws, cell_text(cells[0]) if cells else ''))
        # the ⑦ row specifically
        cells = row_cells(target_row)
        print('  ⑦ ROW: ncells=%d  tcW=%s' % (len(cells), [cell_tcw(c) for c in cells]))
        print('  ⑦ cell text=%r' % cell_text(cells[0], 60))

    # bucket ALL single-cell rows by structural signature
    print('\n===== SINGLE-CELL ROW landscape (the 1323 vs ⑦) =====')
    bucket = Counter()
    bucket_examples = defaultdict(list)
    target_key = None
    for tbl in all_tbls:
        sig = tbl_signature(tbl)
        for tr in tbl.findall(q('tr')):
            cells = row_cells(tr)
            if len(cells) != 1:
                continue
            tcw = cell_tcw(cells[0])
            # structural key: what distinguishes tables
            key = (
                'nrows>=2' if sig['nrows'] >= 2 else 'nrows=1',
                'tblw=%s' % (sig['tblw'][0],),
                'layout=%s' % (sig['layout'],),
                'cellmar' if sig['has_cellmar'] else 'no-cellmar',
                'tcW=%s' % (tcw[0],),
                'style=%s' % (sig['style'],),
            )
            bucket[key] += 1
            if len(bucket_examples[key]) < 2:
                bucket_examples[key].append(cell_text(cells[0], 24))
            if tr is target_row:
                target_key = key

    print('  ⑦ row structural key =', target_key)
    print('\n  bucket counts (sorted desc):')
    for key, n in bucket.most_common():
        mark = '  <== ⑦ HERE' if key == target_key else ''
        print('    %5d  %s%s' % (n, key, mark))
        for ex in bucket_examples[key]:
            print('            ex: %r' % ex)


if __name__ == '__main__':
    main()
