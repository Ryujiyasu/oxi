"""Compare Oxi layout dump vs Word DML for any doc; find first drift per page.

Usage:
  python tools/metrics/diff_layout_vs_word.py <docname-prefix>

Requires:
  - pipeline_data/_<prefix>_layout.json (from oxi-gdi-renderer --dump-layout)
  - C:/Users/ryuji/oxi-main/pipeline_data/word_dml/<full-docname>.json

Reports per-page:
  - Cumulative drift (Oxi - Word) up to each Word paragraph
  - First paragraph where drift jumps > 5pt (likely a bug location)
  - Whether the jump is inside a table
"""
import sys, json, os, glob, re

if len(sys.argv) < 2:
    print("Usage: python diff_layout_vs_word.py <docname-prefix>")
    sys.exit(1)

prefix = sys.argv[1]

# Locate Word DML
dml_files = glob.glob(f"C:/Users/ryuji/oxi-main/pipeline_data/word_dml/{prefix}*.json")
if not dml_files:
    print(f"No Word DML for prefix '{prefix}'")
    sys.exit(1)
dml_path = dml_files[0]
docname = os.path.splitext(os.path.basename(dml_path))[0]
print(f"Word DML: {dml_path}")
print(f"Full docname: {docname}")

# Locate Oxi layout dump
layout_path = f"pipeline_data/_{prefix}_layout.json"
if not os.path.exists(layout_path):
    # Try short name
    layout_path = f"pipeline_data/_{docname[:20]}_layout.json"
    if not os.path.exists(layout_path):
        print(f"Oxi layout dump not found. Generate via:")
        print(f"  tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe \\")
        print(f"    tools/golden-test/documents/docx/{docname}.docx \\")
        print(f"    /tmp/_dummy 150 --dump-layout={layout_path}")
        sys.exit(1)
print(f"Oxi layout: {layout_path}")

wd = json.load(open(dml_path, encoding='utf-8'))
od = json.load(open(layout_path, encoding='utf-8'), strict=False)

# Group Oxi elements by page and y
from collections import defaultdict
print(f"\nWord pages: {wd.get('pages')}, Oxi pages: {len(od['pages'])}")

# Per-Word-page drift analysis
word_pages = defaultdict(list)
for p in wd['paragraphs']:
    word_pages[p['page']].append(p)

for page_num in sorted(word_pages.keys())[:3]:
    word_paras = sorted(word_pages[page_num], key=lambda p: p['index'])
    # Group Word paras by y (table rows share y). Use the MIN index per y-group.
    y_groups = defaultdict(list)
    for p in word_paras:
        y_groups[round(p['y'], 1)].append(p)
    # One representative y per unique-y group
    word_unique_ys = sorted(y_groups.keys())
    # Oxi page
    if page_num - 1 >= len(od['pages']):
        print(f"\np{page_num}: Oxi has no matching page")
        continue
    oxi_page = od['pages'][page_num - 1]
    oxi_text = [e for e in oxi_page['elements'] if e['kind'] == 'text']
    oxi_ys = sorted(set(round(e['y'], 1) for e in oxi_text))
    print(f"\n=== page {page_num}: Word {len(word_paras)} paras ({len(word_unique_ys)} unique y), Oxi {len(oxi_ys)} unique y ===")
    print(f"{'wi':>4} {'wy':>7} {'fs':>5} {'oy':>7} {'dy':>7} {'jump':>6}")
    prev_dy = None
    max_jump = 0
    max_jump_at = None
    n = min(len(word_unique_ys), len(oxi_ys), 30)
    for i in range(n):
        wy = word_unique_ys[i]
        oy = oxi_ys[i]
        # Representative Word para at this y
        rep = y_groups[wy][0]
        fs = rep['font_size']
        dy = oy - wy
        jump = "" if prev_dy is None else f"{dy - prev_dy:+.2f}"
        mark = ""
        if prev_dy is not None and abs(dy - prev_dy) > max_jump:
            max_jump = abs(dy - prev_dy)
            max_jump_at = (i, rep['index'], wy, oy)
        if prev_dy is not None and abs(dy - prev_dy) > 5:
            mark = " ← BIG JUMP"
        n_in_row = len(y_groups[wy])
        row_info = f" (row has {n_in_row} paras)" if n_in_row > 1 else ""
        print(f"{rep['index']:>4} {wy:>7.1f} {fs:>5.1f} {oy:>7.2f} {dy:>+7.2f} {jump:>6}{mark}{row_info}")
        prev_dy = dy
    if max_jump_at:
        i, wi, wy, oy = max_jump_at
        print(f"  Max jump: {max_jump:.2f}pt at word idx {wi} (wy={wy} oy={oy})")
