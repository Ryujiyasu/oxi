"""Compare Oxi p.2 paragraph-level Y positions vs Word p.2 for 0e7a.

If per-para Y matches Word within a few pt, layout IS correct and the
SSIM gap is sub-pixel / AA rendering (layout ceiling per prior memos).

If per-para Y differs significantly, there's a concrete layout drift
to investigate via additive-primary scratch repro.
"""
import json
from collections import defaultdict

OXI = json.load(open('pipeline_data/_0e7a_layout.json', encoding='utf-8'))
WORD = json.load(open('pipeline_data/0e7a_word_paras.json', encoding='utf-8'))

# Oxi p2 elements grouped by para_idx
p2 = OXI['pages'][1]['elements']
oxi_by_pi = defaultdict(list)
for e in p2:
    if e.get('type') != 'text': continue
    pi = e.get('para_idx')
    if pi is None: continue
    oxi_by_pi[pi].append(e)

# First Y per para
oxi_para_y = {pi: min(e['y'] for e in elems) for pi, elems in oxi_by_pi.items()}
sorted_oxi_pi = sorted(oxi_para_y.keys())
print(f'Oxi p2: {len(sorted_oxi_pi)} paragraphs, idx range {sorted_oxi_pi[0]}..{sorted_oxi_pi[-1]}')

# Word p2 paras
word_p2 = sorted([p for p in WORD if p.get('page') == 2], key=lambda p: p.get('y', 0))
print(f'Word p2: {len(word_p2)} paragraphs')

# Attempt alignment: assume both are in order; find common first-chars sequence
# Approximate: list both side-by-side
print(f'\\n{"i":3} {"W_idx":5} {"W_y":>7} {"text":20} | {"O_idx":5} {"O_y":>7}')
print('-' * 80)

# Take first 15 Word paras and try to align with Oxi
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
for i in range(min(20, len(word_p2), len(sorted_oxi_pi))):
    w = word_p2[i] if i < len(word_p2) else None
    o_pi = sorted_oxi_pi[i] if i < len(sorted_oxi_pi) else None
    w_y = w['y'] if w else None
    w_idx = w['idx'] if w else None
    w_text = w.get('text_utf8', '')[:20] if w else ''
    o_y = oxi_para_y.get(o_pi, 0) if o_pi is not None else None
    delta = (o_y - w_y) if (w_y and o_y) else None
    delta_s = f'{delta:+.2f}' if delta is not None else '-'
    print(f'{i:3} {w_idx or 0:5} {w_y:7.2f} {w_text:20} | {o_pi or 0:5} {o_y:7.2f} {delta_s}')
