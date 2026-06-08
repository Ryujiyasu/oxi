# -*- coding: utf-8 -*-
"""List every paragraph with jc=center or jc=right (or w:ind) in a docx, with its
concatenated text, alignment, indent, snapToGrid, and whether it's inside a table cell.
cp932-safe: UTF-8 read, results to a file. Helps see what the S502-touched lines really are."""
import sys, io, zipfile, re

docx, out = sys.argv[1], sys.argv[2]
xml = zipfile.ZipFile(docx).read('word/document.xml').decode('utf-8')

# crude cell-depth tracking by scanning tags in order
L = []
# split into paragraphs
for m in re.finditer(r'<w:p\b[^>]*>(.*?)</w:p>', xml, re.S):
    para = m.group(0)
    body = m.group(1)
    ppr = re.search(r'<w:pPr>(.*?)</w:pPr>', body, re.S)
    if not ppr:
        continue
    pprx = ppr.group(1)
    jc = re.search(r'<w:jc w:val="([^"]+)"', pprx)
    jcv = jc.group(1) if jc else ''
    if jcv not in ('center', 'right', 'distribute'):
        continue
    ind = re.search(r'<w:ind ([^/]*)/>', pprx)
    indv = ind.group(1).strip() if ind else ''
    sng = 'snapToGrid' in pprx
    sng0 = re.search(r'<w:snapToGrid w:val="0"', pprx) is not None
    # text
    txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', body, re.S))
    # in cell? count tc open/close before this para's position
    pos = m.start()
    in_cell = xml.count('<w:tc>', 0, pos) > xml.count('</w:tc>', 0, pos)
    L.append('jc=%-10s cell=%-5s snapGrid0=%-5s ind=[%s]  text=%s' % (
        jcv, in_cell, sng0, indv, txt[:30]))

with io.open(out, 'w', encoding='utf-8') as f:
    f.write('count=%d\n' % len(L))
    f.write('\n'.join(L) + '\n')
print('wrote', out, 'count', len(L))
