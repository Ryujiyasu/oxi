# -*- coding: utf-8 -*-
"""Diagnostic: dump Word EMF lines and Oxi dump lines side-by-side (geometry + text) to JSON.
cp932-safe: read the JSON, never eyeball console Japanese. Helps design line matching."""
import json, sys
SC = 0.12


def cluster(glyphs, tol):
    gs = sorted(glyphs, key=lambda g: (round(g['y'], 1), g['x']))
    lines = []
    for g in gs:
        if lines and abs(g['y'] - lines[-1]['y']) < tol:
            lines[-1]['gs'].append(g)
        else:
            lines.append({'y': g['y'], 'gs': [g]})
    for L in lines:
        L['gs'].sort(key=lambda g: g['x'])
        L['y'] = sum(g['y'] for g in L['gs']) / len(L['gs'])
    return lines


emf_path, dump_path, pidx, out = sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4]
W = json.load(open(emf_path, encoding='utf-8'))['records']
O = json.load(open(dump_path, encoding='utf-8'))['pages'][pidx]['elements']

wg = []
for r in W:
    if not r['text'].strip():
        continue
    x = r['x']
    dx = r.get('dx', [])
    for k, ch in enumerate(r['text']):
        wg.append({'char': ch, 'x': x * SC, 'y': r['y'] * SC})
        x += dx[k] if k < len(dx) else 75
og = [{'char': e['text'], 'x': e['x'], 'y': e['y'], 'w': e.get('w', 0)}
      for e in O if e.get('type') == 'text' and e.get('text', '').strip()]

wlines = cluster(wg, 5.0)
olines = cluster(og, 5.0)


def lineinfo(L):
    gs = L['gs']
    return {'y': round(L['y'], 1), 'x0': round(gs[0]['x'], 1), 'x1': round(gs[-1]['x'], 1),
            'n': len(gs), 'text': ''.join(g['char'] for g in gs)}


res = {'word': [lineinfo(L) for L in wlines], 'oxi': [lineinfo(L) for L in olines]}
json.dump(res, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
print("word_lines=%d oxi_lines=%d -> %s" % (len(wlines), len(olines), out))
print("word total chars=%d  oxi total chars=%d" % (sum(L['n'] for L in wlines), sum(L['n'] for L in olines)))
