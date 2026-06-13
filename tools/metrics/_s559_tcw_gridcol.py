# -*- coding: utf-8 -*-
"""S559 — dump tcW vs gridCol for every cell in the ⑦-signature bucket
(nrows=1, tblw=auto, no-cellmar single-cell tables). The structural signature
is identical across 87 cells, so the discriminator (why Word reserves cellMar
for ⑦ but the universal-fire regressed) must be in the NUMERIC tcW vs gridCol
relationship. Hypothesis: ⑦ has tcW - gridCol ~= cellMar (216tw); others have
tcW == gridCol. Print the distribution and flag ⑦.
"""
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

DOCX = r'c:\tmp\3a4f9f.docx'
ANCHOR = u'常に整理整頓'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
sys.stdout.reconfigure(encoding='utf-8')


def q(tag):
    return W + tag


def attr(el, name):
    return el.get(W + name) if el is not None else None


def cell_text(tc, limit=30):
    return ''.join(t.text or '' for t in tc.iter(q('t')))[:limit]


def main():
    with zipfile.ZipFile(DOCX) as z:
        root = ET.fromstring(z.read('word/document.xml'))

    rows = []  # (tcW_int, gridcol_int, diff, is_target, npara, text)
    diff_counter = Counter()
    for tbl in root.iter(q('tbl')):
        pr = tbl.find(q('tblPr'))
        tblw_el = pr.find(q('tblW')) if pr is not None else None
        tblw_type = attr(tblw_el, 'type') if tblw_el is not None else None
        has_cm = pr is not None and pr.find(q('tblCellMar')) is not None
        g = tbl.find(q('tblGrid'))
        gcols = [attr(c, 'w') for c in g.findall(q('gridCol'))] if g is not None else []
        trs = tbl.findall(q('tr'))
        if len(trs) != 1 or len(gcols) != 1:
            continue
        # single-row single-col table = ⑦ bucket family
        if tblw_type != 'auto' or has_cm:
            continue
        tr = trs[0]
        cells = tr.findall(q('tc'))
        if len(cells) != 1:
            continue
        tc = cells[0]
        tcpr = tc.find(q('tcPr'))
        tcw_el = tcpr.find(q('tcW')) if tcpr is not None else None
        tcw = attr(tcw_el, 'w') if tcw_el is not None else None
        try:
            tcw_i = int(tcw)
            gc_i = int(gcols[0])
        except (TypeError, ValueError):
            continue
        diff = tcw_i - gc_i
        npara = len(tc.findall(q('p')))
        is_target = ANCHOR in ''.join(t.text or '' for t in tc.iter(q('t')))
        rows.append((tcw_i, gc_i, diff, is_target, npara, cell_text(tc)))
        diff_counter[diff] += 1

    print('⑦-family single-cell tables:', len(rows))
    print('\ndiff (tcW - gridCol) distribution:')
    for d, n in diff_counter.most_common():
        print('   diff=%5d  count=%3d' % (d, n))

    print('\nrows where ⑦ lives + a few examples per diff value:')
    seen_diff = set()
    total_para = 0
    for tcw_i, gc_i, diff, is_target, npara, text in rows:
        total_para += npara
        show = is_target or diff not in seen_diff
        seen_diff.add(diff)
        if show:
            mark = '  <== ⑦' if is_target else ''
            print('   tcW=%5d gridCol=%5d diff=%5d npara=%3d%s  %r'
                  % (tcw_i, gc_i, diff, npara, mark, text))
    print('\ntotal paragraphs across these cells:', total_para)

    # split paragraph counts by diff sign
    para_by_diff = Counter()
    for tcw_i, gc_i, diff, is_target, npara, text in rows:
        para_by_diff[diff] += npara
    print('\nparagraphs by diff value:')
    for d, n in para_by_diff.most_common():
        print('   diff=%5d  paras=%4d' % (d, n))


if __name__ == '__main__':
    main()
