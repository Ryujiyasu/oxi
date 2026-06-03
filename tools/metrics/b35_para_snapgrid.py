# -*- coding: utf-8 -*-
"""Parse b35 docx table0: per cell, per paragraph -> font size (sz/2 pt), snapToGrid (0 if
<w:snapToGrid w:val="0"/> present in pPr, else 1=default), spacing before/after/line. Goal:
test whether the 9pt lines that Oxi places +3pt too low are exactly the snapToGrid=0 paras.
cp932-safe (JSON out, ASCII console)."""
import zipfile, json, sys
import xml.etree.ElementTree as ET

docx = sys.argv[1] if len(sys.argv) > 1 else \
    'tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
root = ET.fromstring(zipfile.ZipFile(docx).read('word/document.xml'))
body = root.find(W + 'body')
tbl = body.findall('.//' + W + 'tbl')[0]

out = []
for ri, tr in enumerate(tbl.findall(W + 'tr')):
    for ci, tc in enumerate(tr.findall(W + 'tc')):
        for pi, p in enumerate(tc.findall(W + 'p')):
            pPr = p.find(W + 'pPr')
            snap = 1
            sz = None
            spb = spa = spline = splrule = None
            if pPr is not None:
                sg = pPr.find(W + 'snapToGrid')
                if sg is not None and sg.get(W + 'val') in ('0', 'false'):
                    snap = 0
                sp = pPr.find(W + 'spacing')
                if sp is not None:
                    spb = sp.get(W + 'before'); spa = sp.get(W + 'after')
                    spline = sp.get(W + 'line'); splrule = sp.get(W + 'lineRule')
                rpr = pPr.find(W + 'rPr')
                if rpr is not None:
                    szel = rpr.find(W + 'sz')
                    if szel is not None:
                        sz = int(szel.get(W + 'val')) / 2
            # fall back: first run sz
            if sz is None:
                for r in p.findall(W + 'r'):
                    rpr = r.find(W + 'rPr')
                    if rpr is not None:
                        szel = rpr.find(W + 'sz')
                        if szel is not None:
                            sz = int(szel.get(W + 'val')) / 2; break
            txt = ''.join((t.text or '') for t in p.iter(W + 't'))
            out.append({'row': ri, 'col': ci, 'para': pi, 'sz': sz, 'snap': snap,
                        'spb': spb, 'spa': spa, 'line': spline, 'lineRule': splrule,
                        'text': txt[:14]})

json.dump(out, open(r'c:/tmp/b35_para_snap.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
# summarize: col1 paras with sz/snap, flag the +3-jump rows
jump_rows = {1, 3, 4, 8, 9, 12}
print('COL1 paragraphs (row: para sz/snap/spacing):')
for c in out:
    if c['col'] != 1:
        continue
    mark = '  <JUMP row' if c['row'] in jump_rows else ''
    print('  row%2d p%d sz=%s snap=%d spb=%s spa=%s line=%s/%s%s' % (
        c['row'], c['para'], c['sz'], c['snap'], c['spb'], c['spa'], c['line'], c['lineRule'], mark))
import collections
print('\nsnap=0 count among col1:', sum(1 for c in out if c['col'] == 1 and c['snap'] == 0))
print('snap by row (col1, any para snap=0?):')
byrow = collections.defaultdict(lambda: {'has9': False, 'snap0': False})
for c in out:
    if c['col'] != 1:
        continue
    if c['sz'] == 9.0:
        byrow[c['row']]['has9'] = True
        if c['snap'] == 0:
            byrow[c['row']]['snap0'] = True
for r in sorted(byrow):
    j = 'JUMP' if r in jump_rows else 'flat'
    print('  row%2d has9pt=%s snap0=%s  measured=%s' % (r, byrow[r]['has9'], byrow[r]['snap0'], j))
