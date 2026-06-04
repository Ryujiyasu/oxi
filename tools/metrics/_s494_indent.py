# -*- coding: utf-8 -*-
"""Dump paragraph indent attributes for the paragraphs that produce text on a page.
ASCII-only output. Helps characterize a horizontal indent offset (chars vs twip, hanging,
firstLine). cp932-safe."""
import sys, zipfile, re

path = sys.argv[1]
z = zipfile.ZipFile(path)
doc = z.read('word/document.xml').decode('utf-8', 'replace')
body = doc.split('<w:body>', 1)[-1]
count = 0
for m in re.finditer(r'<w:p[ >].*?</w:p>', body, re.S):
    seg = m.group(0)
    if '<w:t>' not in seg and '<w:t ' not in seg:
        continue
    ind = re.search(r'<w:ind([^>]*)/>', seg)
    attrs = {}
    if ind:
        for k in ['left', 'start', 'right', 'firstLine', 'hanging',
                  'leftChars', 'startChars', 'firstLineChars', 'hangingChars']:
            mm = re.search(r'w:%s="(-?\d+)"' % k, ind.group(1))
            if mm:
                attrs[k] = mm.group(1)
    sz = re.search(r'<w:sz w:val="(\d+)"', seg)
    jc = re.search(r'<w:jc w:val="(\w+)"', seg)
    pstyle = re.search(r'<w:pStyle w:val="([^"]+)"', seg)
    sample = ''.join(c for c in re.sub(r'<[^>]+>', '', seg)[:14])
    sample_ascii = ''.join(ch if ord(ch) < 128 else '.' for ch in sample)
    print('p%d sz=%s jc=%s style=%s ind=%s  | %s'
          % (count, sz.group(1) if sz else '-', jc.group(1) if jc else '-',
             pstyle.group(1) if pstyle else '-', attrs or '{}', sample_ascii))
    count += 1
    if count >= 40:
        break
