# -*- coding: utf-8 -*-
"""Extract first-paragraph line-spacing mode + docGrid + first-run font/size for a set
of docx. ASCII-only output (no Japanese), cp932-safe. Helps test whether the first-line
half-leading direction (over vs under reserve) splits by lineRule/docGrid mode."""
import sys, zipfile, re


def first_para_info(path):
    z = zipfile.ZipFile(path)
    doc = z.read('word/document.xml').decode('utf-8', 'replace')
    # docGrid
    dg = re.search(r'<w:docGrid[^>]*w:type="(\w+)"[^>]*?(?:w:linePitch="(\d+)")?', doc)
    dgtype = dg.group(1) if dg else 'none'
    lp = re.search(r'w:linePitch="(\d+)"', dg.group(0)) if dg else None
    linePitch = lp.group(1) if lp else '?'
    # first paragraph spacing
    body = doc.split('<w:body>', 1)[-1]
    pstart = body.find('<w:p ')
    if pstart < 0:
        pstart = body.find('<w:p>')
    # find first few paras
    out = []
    pos = 0
    count = 0
    for m in re.finditer(r'<w:p[ >].*?</w:p>', body, re.S):
        seg = m.group(0)
        if '<w:t>' not in seg and '<w:t ' not in seg:
            continue
        spc = re.search(r'<w:spacing([^>]*)/>', seg)
        line = lineRule = before = beforeLines = '-'
        if spc:
            a = spc.group(1)
            ml = re.search(r'w:line="(-?\d+)"', a); line = ml.group(1) if ml else '-'
            mr = re.search(r'w:lineRule="(\w+)"', a); lineRule = mr.group(1) if mr else '-'
            mb = re.search(r'w:before="(-?\d+)"', a); before = mb.group(1) if mb else '-'
            mbl = re.search(r'w:beforeLines="(-?\d+)"', a); beforeLines = mbl.group(1) if mbl else '-'
        sz = re.search(r'<w:sz w:val="(\d+)"', seg)
        sznum = sz.group(1) if sz else '-'
        snap = 'snap0' if 'w:snapToGrid w:val="0"' in seg else 'snap1'
        out.append((line, lineRule, before, beforeLines, sznum, snap))
        count += 1
        if count >= 3:
            break
    print('  docGrid=%s linePitch=%s' % (dgtype, linePitch))
    for i, (l, lr, b, bl, s, sn) in enumerate(out):
        print('    para%d: line=%s lineRule=%s before=%s beforeLines=%s sz=%s %s' % (i, l, lr, b, bl, s, sn))


if __name__ == '__main__':
    for p in sys.argv[1:]:
        import os
        print('===', os.path.basename(p), '===')
        try:
            first_para_info(p)
        except Exception as e:
            print('  ERR', e)
