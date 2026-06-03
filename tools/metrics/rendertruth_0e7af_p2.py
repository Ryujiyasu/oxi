# -*- coding: utf-8 -*-
"""S492 big-job FIRST CLEAN application: per-line render-truth dy for 0e7af p2 (the cleanest
target: 0 textboxes, docGrid=none, position-capped). Group Word EMF records + Oxi dump elements
into lines (by baseline / line-box-top), order-match (no overlap so order is reliable), compute
per-line dy = Word_baseline - Oxi_baseline. Uniform = body-top/offset; growing = line-height
accumulation; per-line steps = per-para. UTF-8 file (cp932 rule); ASCII verdict + JSON out."""
import json
SC = 0.12
W = json.load(open(r'c:\tmp\0e7af_emf_p2_pos.json', encoding='utf-8'))['records']
O = json.load(open(r'c:\tmp\0e7af_dump.json', encoding='utf-8'))['pages'][1]['elements']

# Word EMF lines: cluster records by baseline (raw*SC)
wruns = sorted([(r['y'] * SC, r['x'] * SC, r['text']) for r in W if r['text'].strip()])
wlines = []
for y, x, t in wruns:
    if wlines and abs(y - wlines[-1]['base']) < 3:
        wlines[-1]['runs'].append((x, t))
    else:
        wlines.append({'base': y, 'runs': [(x, t)]})
for L in wlines:
    L['runs'].sort()
    L['text'] = ''.join(t for _, t in L['runs'])
    L['x0'] = L['runs'][0][0]

# Oxi lines: cluster text els by line-box top; baseline = top + text_y_off + 0.8*fs
oels = [e for e in O if e.get('type') == 'text' and e.get('text', '').strip()]
oels.sort(key=lambda e: (round(e['y'], 1), e['x']))
olines = []
for e in oels:
    top = e['y']
    if olines and abs(top - olines[-1]['top']) < 3:
        olines[-1]['els'].append(e)
    else:
        olines.append({'top': top, 'els': [e]})
for L in olines:
    L['els'].sort(key=lambda e: e['x'])
    L['text'] = ''.join(e['text'] for e in L['els'])
    L['fs'] = max(e['font_size'] for e in L['els'])
    L['tyoff'] = L['els'][0].get('text_y_off', 0)
    L['base'] = L['top'] + L['tyoff'] + 0.8 * L['fs']
    L['x0'] = L['els'][0]['x']

print("0e7af p2: Word EMF lines=%d  Oxi lines=%d" % (len(wlines), len(olines)))
# order-match (both sorted by y); align by count
n = min(len(wlines), len(olines))
pairs = []
for i in range(n):
    w, o = wlines[i], olines[i]
    pairs.append({'i': i, 'wbase': round(w['base'], 1), 'obase': round(o['base'], 1),
                  'dy': round(w['base'] - o['base'], 1), 'wx': round(w['x0'], 1), 'ox': round(o['x0'], 1),
                  'fs': o['fs'], 'wn': len(w['text'].strip()), 'on': len(o['text'].strip())})
json.dump(pairs, open(r'c:\tmp\rt_0e7af_p2.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print("per-line dy (Word_base - Oxi_base); x0 to sanity-check alignment:")
for p in pairs:
    flag = '' if abs(p['wx'] - p['ox']) < 15 else '  <x-mismatch (maybe misaligned)'
    print("  L%2d wbase=%.1f obase=%.1f dy=%+.1f  wx=%.0f ox=%.0f fs=%.0f wn=%d on=%d%s" % (
        p['i'], p['wbase'], p['obase'], p['dy'], p['wx'], p['ox'], p['fs'], p['wn'], p['on'], flag))
import statistics
dys = [p['dy'] for p in pairs if abs(p['wx'] - p['ox']) < 15]
if dys:
    print("\naligned-line dy: n=%d mean=%+.2f median=%+.2f first=%+.1f last=%+.1f spread=%.1f" % (
        len(dys), statistics.mean(dys), statistics.median(dys), dys[0], dys[-1], max(dys) - min(dys)))
    print("UNIFORM=body-Y offset; GROWING(first->last)=line-height accumulation; STEPS=per-para")
