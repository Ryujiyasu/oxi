"""Compare b35 p.1 per-paragraph Y — Oxi vs Word."""
import json, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

OXI = json.load(open('pipeline_data/_b35_layout.json', encoding='utf-8'))
WORD = json.load(open('pipeline_data/b35_word_paras.json', encoding='utf-8'))

# Oxi p1 elements
p1 = OXI['pages'][0]['elements']
oxi_by_pi = defaultdict(list)
for e in p1:
    if e.get('type') == 'text' and e.get('para_idx') is not None:
        oxi_by_pi[e['para_idx']].append(e)
oxi_paras = []
for pi in sorted(oxi_by_pi):
    elems = sorted(oxi_by_pi[pi], key=lambda e: (e['y'], e['x']))
    first_y = min(e['y'] for e in elems)
    first_text = ''.join(e.get('text', '') for e in elems if abs(e['y'] - first_y) < 0.5)
    oxi_paras.append({'pi': pi, 'y': first_y, 'text': first_text[:40]})

print(f'Oxi b35 p.1: {len(oxi_paras)} paragraphs')
word_paras = sorted([p for p in WORD if p.get('page') == 1], key=lambda p: p.get('y', 0))
print(f'Word b35 p.1: {len(word_paras)} paragraphs')

# Align by first few chars of text
aligned = []
used_oxi = set()
for w in word_paras:
    w_text = w.get('text_utf8', '').strip()
    if len(w_text) < 3: continue
    prefix = w_text[:5]
    for o in oxi_paras:
        if o['pi'] in used_oxi: continue
        if prefix[:3] and prefix[:3] in o['text'][:10]:
            aligned.append((w, o))
            used_oxi.add(o['pi'])
            break

print(f'\\nAligned: {len(aligned)} pairs')
print(f'{"w_idx":>5} {"W_y":>7} {"O_y":>7} {"Δ":>7}  text')
deltas = []
for w, o in aligned:
    delta = o['y'] - w['y']
    deltas.append(delta)
    print(f"{w.get('idx', 0):5} {w['y']:7.2f} {o['y']:7.2f} {delta:+7.2f}  {w.get('text_utf8', '')[:35]}")

if deltas:
    import statistics
    print(f'\\n|Δ|: median={statistics.median(abs(d) for d in deltas):.2f} max={max(abs(d) for d in deltas):.2f}')
    print(f'signed Δ: median={statistics.median(deltas):+.2f} min={min(deltas):+.2f} max={max(deltas):+.2f}')
