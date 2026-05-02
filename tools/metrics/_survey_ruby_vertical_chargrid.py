# -*- coding: utf-8 -*-
"""Survey baseline for ruby/vertical/charGrid usage to identify investigation targets."""
import sys, os, glob, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX_DIR = r'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx'

results = []
for path in sorted(glob.glob(os.path.join(DOCX_DIR, '*.docx'))):
    name = os.path.basename(path)
    feature_counts = {
        'ruby': 0,                # <w:ruby>
        'vert_textdir': 0,        # <w:textDirection w:val="tb..">
        'vert_eavert': 0,         # eaVert
        'vert_pagedir': 0,        # <w:bidi/> (RTL mostly)
        'no_grid': 0,             # docGrid type="default"
        'lines_grid': 0,          # docGrid type="lines"
        'linesAndChars_grid': 0,  # docGrid type="linesAndChars"
        'snap_to_grid_off': 0,    # snapToGrid=0
    }
    try:
        with zipfile.ZipFile(path) as z:
            try:
                doc = z.read('word/document.xml').decode('utf-8', errors='replace')
                feature_counts['ruby'] = doc.count('<w:ruby>')
                feature_counts['vert_textdir'] = len(re.findall(r'<w:textDirection[^/>]*?w:val="tb', doc))
                feature_counts['vert_eavert'] = doc.count('eaVert')
                feature_counts['vert_pagedir'] = doc.count('<w:bidi/>')
                feature_counts['snap_to_grid_off'] = len(re.findall(r'<w:snapToGrid[^/>]*?w:val="0"', doc))
                # docGrid: in sectPr
                if 'w:type="default"' in doc:
                    feature_counts['no_grid'] = doc.count('<w:docGrid w:type="default"')
                if 'w:type="lines"' in doc:
                    feature_counts['lines_grid'] = doc.count('<w:docGrid w:type="lines"')
                if 'w:type="linesAndChars"' in doc:
                    feature_counts['linesAndChars_grid'] = doc.count('<w:docGrid w:type="linesAndChars"')
            except KeyError:
                continue
    except Exception:
        continue
    if any(feature_counts.values()):
        results.append((name, feature_counts))

print(f"=== Ruby usage ===")
ruby_docs = [(n, c) for n, c in results if c['ruby'] > 0]
print(f"Docs with <w:ruby>: {len(ruby_docs)}")
for n, c in ruby_docs[:15]:
    print(f"  {n}: {c['ruby']} ruby instances")

print(f"\n=== Vertical writing ===")
vert_docs = [(n, c) for n, c in results if c['vert_textdir'] > 0 or c['vert_eavert'] > 0]
print(f"Docs with vertical writing: {len(vert_docs)}")
for n, c in vert_docs[:10]:
    print(f"  {n}: textDir={c['vert_textdir']} eaVert={c['vert_eavert']}")

print(f"\n=== docGrid type distribution ===")
no_grid = sum(1 for n, c in results if c['no_grid'] > 0)
lines_only = sum(1 for n, c in results if c['lines_grid'] > 0 and c['linesAndChars_grid'] == 0)
linesChars = sum(1 for n, c in results if c['linesAndChars_grid'] > 0)
print(f"  no_grid (type=default): {no_grid}")
print(f"  lines_grid: {lines_only}")
print(f"  linesAndChars_grid: {linesChars}")
