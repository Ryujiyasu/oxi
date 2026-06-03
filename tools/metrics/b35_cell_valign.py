# -*- coding: utf-8 -*-
"""Parse b35 docx table: per (row,col) cell vAlign + trHeight + n paragraphs. Output JSON to
correlate with the measured per-cell vertical dy. cp932-safe (results to JSON, ASCII console)."""
import zipfile, json, sys, re
import xml.etree.ElementTree as ET

docx = sys.argv[1] if len(sys.argv) > 1 else \
    'tools/golden-test/documents/docx/b35123fe8efc_tokumei_08_01.docx'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
xml = zipfile.ZipFile(docx).read('word/document.xml')
root = ET.fromstring(xml)
body = root.find(W + 'body')

tables = body.findall('.//' + W + 'tbl')
out = []
for ti, tbl in enumerate(tables):
    rows = tbl.findall(W + 'tr')
    for ri, tr in enumerate(rows):
        # row height
        trPr = tr.find(W + 'trPr')
        trh = None; trhrule = None
        if trPr is not None:
            h = trPr.find(W + 'trHeight')
            if h is not None:
                trh = h.get(W + 'val'); trhrule = h.get(W + 'hRule')
        cells = tr.findall(W + 'tc')
        for ci, tc in enumerate(cells):
            tcPr = tc.find(W + 'tcPr')
            valign = None
            if tcPr is not None:
                va = tcPr.find(W + 'vAlign')
                if va is not None:
                    valign = va.get(W + 'val')
            paras = tc.findall(W + 'p')
            # text content (first 12 chars) for matching
            txt = ''
            for p in paras:
                for t in p.iter(W + 't'):
                    txt += (t.text or '')
            out.append({'tbl': ti, 'row': ri, 'col': ci, 'valign': valign,
                        'trHeight': trh, 'trHeightRule': trhrule,
                        'n_para': len(paras), 'text': txt[:16]})

json.dump(out, open(r'c:/tmp/b35_cell_valign.json', 'w', encoding='utf-8'),
          ensure_ascii=False, indent=0)
import collections
vc = collections.Counter(c['valign'] for c in out)
print('tables=%d  total cells=%d' % (len(tables), len(out)))
print('vAlign distribution:', dict(vc))
print('trHeightRule distribution:', dict(collections.Counter(c['trHeightRule'] for c in out)))
# per (row,col) valign for table 0
print('table0 per-cell valign (row,col -> valign, npara, trHeight):')
for c in out:
    if c['tbl'] != 0:
        continue
    print('  (%d,%d) valign=%-6s npara=%d trH=%s/%s' % (
        c['row'], c['col'], str(c['valign']), c['n_para'], c['trHeight'], c['trHeightRule']))
