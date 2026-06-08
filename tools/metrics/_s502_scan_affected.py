# -*- coding: utf-8 -*-
"""S502 affected-set scan: list docs that have docGrid linesAndChars with POSITIVE
charSpace AND at least one jc=center paragraph inside a table cell (the only config S502
changes). Fast (XML only, no render). cp932-safe. Flags bottom-N docs."""
import os, glob, io, re, zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
import json
BASE = json.load(io.open(os.path.join(ROOT, 'pipeline_data', 'ssim_baseline.json'), encoding='utf-8'))
bottom = set(k for k, _ in sorted(((sum(v.values()) / len(v), k) for k, v in BASE.items()))[:12])

dirs = ['tools/golden-test/documents/docx', 'pipeline_data/docx']
seen = set()
hits = []
for d in dirs:
    for p in glob.glob(os.path.join(ROOT, d, '*.docx')):
        name = os.path.splitext(os.path.basename(p))[0]
        if name in seen:
            continue
        seen.add(name)
        try:
            xml = zipfile.ZipFile(p).read('word/document.xml').decode('utf-8', 'ignore')
        except Exception:
            continue
        m = re.search(r'<w:docGrid[^>]*w:type="linesAndChars"[^>]*w:charSpace="(-?\d+)"', xml)
        if not m:
            m2 = re.search(r'<w:docGrid[^>]*w:charSpace="(-?\d+)"[^>]*w:type="linesAndChars"', xml)
            cs = int(m2.group(1)) if m2 else None
        else:
            cs = int(m.group(1))
        if cs is None or cs <= 0:
            continue
        # any jc=center inside a cell?
        center = False
        for pm in re.finditer(r'<w:p\b[^>]*>(.*?)</w:p>', xml, re.S):
            if '<w:jc w:val="center"' in pm.group(1):
                pos = pm.start()
                if xml.count('<w:tc>', 0, pos) > xml.count('</w:tc>', 0, pos):
                    center = True
                    break
        if center:
            base_key = next((k for k in BASE if name.startswith(k) or k.startswith(name) or k == name), name)
            hits.append((name, cs, name in bottom or base_key in bottom))

out = ['S502 affected docs (linesAndChars + charSpace>0 + jc=center in cell): %d' % len(hits)]
for name, cs, isbot in sorted(hits, key=lambda h: (not h[2], h[0])):
    out.append('  %-44s charSpace=%-6d %s' % (name[:44], cs, 'BOTTOM-N' if isbot else ''))
with io.open('c:/tmp/_s502_affected_out.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(out) + '\n')
print('\n'.join(out[:40]))
print('wrote c:/tmp/_s502_affected_out.txt')
