"""Compare Oxi vs Word body-text line positions for b837 p.1 specifically.
Focus: find the first line where they diverge (which paragraph adds drift).
"""
import json, os
from collections import defaultdict

OXI_PATH = os.environ["TMP"] + r"\b837_layout_fresh.txt"
DML_PATH = r"pipeline_data/word_dml/b837808d0555_20240705_resources_data_guideline_02.json"

# Parse Oxi: get body-text glyph clusters per page, excluding footnote area.
# Footnote area starts at ~y=685 for b837 (observed).
FN_Y_THRESHOLD = 680.0

def oxi_body_lines(page_target):
    records = []
    cur = 0
    with open(OXI_PATH, encoding='utf-8') as f:
        for line in f:
            line = line.rstrip()
            if line.startswith('PAGE'):
                cur = int(line.split('\t')[1]) + 1
            elif cur == page_target and line.startswith('TEXT'):
                parts = line.split('\t')
                try:
                    x = float(parts[1]); y = float(parts[2])
                    if y < FN_Y_THRESHOLD:
                        records.append((y, x, None))
                except ValueError:
                    pass
            elif cur == page_target and line.startswith('T\t') and records:
                ch = line[2:]
                y, x, _ = records[-1]
                records[-1] = (y, x, ch)
    # Cluster by y with 10pt tol; aggregate text per cluster
    by_y = defaultdict(list)
    for y, x, ch in records:
        if ch:
            by_y[y].append((x, ch))
    # Sort + cluster within 10pt
    sorted_ys = sorted(by_y.keys())
    if not sorted_ys:
        return []
    clusters = [[sorted_ys[0]]]
    for y in sorted_ys[1:]:
        if y - clusters[-1][-1] <= 10.0:
            clusters[-1].append(y)
        else:
            clusters.append([y])
    lines = []
    for cl in clusters:
        all_glyphs = []
        for y in cl:
            all_glyphs.extend(by_y[y])
        all_glyphs.sort(key=lambda g: g[0])
        # Filter out single-digit superscripts (len < 3) at small x positions?
        # For b837 superscript markers: they are 1-2 chars. To reliably
        # separate body from markers, we count GLYPHS. If cluster has < 5
        # glyphs and only 1-2 unique chars, likely a marker. Keep the main
        # baseline y (the cluster min or max depending on offset direction).
        text = "".join(g[1] for g in all_glyphs)
        # Use max y in cluster as baseline (superscripts float above)
        lines.append((max(cl), text[:50]))
    return lines

with open(DML_PATH, encoding='utf-8') as f:
    dml = json.load(f)

def word_body_lines(page_target):
    lines = []
    for p in dml['paragraphs']:
        if p['page'] != page_target:
            continue
        text = p.get('text','')[:50]
        for line in p.get('lines',[]):
            lines.append((line['y'], text))
    # Sort by y
    lines.sort()
    return lines

for page in [1, 2, 3, 4, 5]:
    print(f"\n===== PAGE {page} =====")
    o = oxi_body_lines(page)
    w = word_body_lines(page)
    print(f"Oxi lines: {len(o)}, Word lines: {len(w)}")
    # Zip for side-by-side
    for i in range(max(len(o), len(w))):
        wy = w[i][0] if i < len(w) else None
        oy = o[i][0] if i < len(o) else None
        if wy is None and oy is None:
            continue
        diff = (oy - wy) if (oy is not None and wy is not None) else None
        wt = w[i][1][:30] if i < len(w) else ""
        ot = o[i][1][:30] if i < len(o) else ""
        print(f"  i={i:2d} | Word y={str(wy):>8} {wt:<30} | Oxi y={str(oy):>8}{diff!s:>10} {ot}")
