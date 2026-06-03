# -*- coding: utf-8 -*-
"""S492 big-job: robust LINE-CLUSTER render-truth dy for 1ec1 blk3 box. Group Word EMF + Oxi
into lines (cluster by baseline), concatenate run text per line, EXCLUDE Oxi empty lines, order-
match Word<->Oxi lines, per-line dy = Word_baseline - Oxi_baseline. Uniform dy = box-position/
v-anchor offset; growing dy = line-height accumulation. UTF-8 file (cp932 rule); results to file
+ ASCII verdict (no eyeballing Japanese console)."""
import json
SC = 0.12
W = json.load(open(r'c:\tmp\1ec1_emf_pos.json', encoding='utf-8'))['records']
O = json.load(open(r'c:\tmp\_1ec1.json', encoding='utf-8'))['pages'][0]['elements']
OUT = r'c:\tmp\rt_lines_1ec1.json'

# Word EMF lines in box region (x 40-445pt, baseline 165-300pt)
wruns = []
for r in W:
    if not r['text'].strip():
        continue
    xpt, ypt = r['x'] * SC, r['y'] * SC
    if 40 <= xpt <= 445 and 165 <= ypt <= 300:
        wruns.append((ypt, xpt, r['text']))
wruns.sort()
# cluster by baseline within 3pt
wlines = []
for y, x, t in wruns:
    if wlines and abs(y - wlines[-1]['y']) < 3:
        wlines[-1]['runs'].append((x, t))
    else:
        wlines.append({'y': y, 'runs': [(x, t)]})
for L in wlines:
    L['runs'].sort()
    L['text'] = ''.join(t for _, t in L['runs'])
    L['nchar'] = len(L['text'].strip())

# Oxi lines in box (glyph_top 168-310, x 40-445); baseline = top + 0.8*fs; exclude empty
oruns = []
for e in O:
    if e.get('type') != 'text':
        continue
    txt = e.get('text', '')
    gtop = e['y'] + e.get('text_y_off', 0)
    if 40 <= e['x'] <= 445 and 168 <= gtop <= 310:
        oruns.append((gtop, e['x'], txt, e['font_size']))
oruns.sort()
olines = []
for gt, x, t, fs in oruns:
    if olines and abs(gt - olines[-1]['top']) < 3:
        olines[-1]['runs'].append((x, t, fs))
    else:
        olines.append({'top': gt, 'runs': [(x, t, fs)]})
for L in olines:
    L['runs'].sort()
    L['text'] = ''.join(t for _, t, _ in L['runs'])
    L['nchar'] = len(L['text'].strip())
    L['fs'] = max(fs for _, _, fs in L['runs'])
    L['base'] = L['top'] + 0.8 * L['fs']

o_nonempty = [L for L in olines if L['nchar'] > 0]
o_empty = [L for L in olines if L['nchar'] == 0]

# order-match Word lines to Oxi non-empty lines
pairs = []
for i in range(min(len(wlines), len(o_nonempty))):
    w = wlines[i]; o = o_nonempty[i]
    pairs.append({'i': i, 'w_base': round(w['y'], 1), 'o_base': round(o['base'], 1),
                  'dy': round(w['y'] - o['base'], 1), 'w_nchar': w['nchar'], 'o_nchar': o['nchar'],
                  'w_text': w['text'][:10], 'o_text': o['text'][:10]})

res = {'n_word_lines': len(wlines), 'n_oxi_nonempty': len(o_nonempty),
       'n_oxi_empty': len(o_empty), 'pairs': pairs}
json.dump(res, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

print("1ec1 blk3 box: Word lines=%d  Oxi non-empty lines=%d  Oxi empty lines=%d" % (
    len(wlines), len(o_nonempty), len(o_empty)))
print("order-matched per-line dy (Word_baseline - Oxi_baseline):")
for p in pairs:
    print("  L%d  Word=%.1f Oxi=%.1f  dy=%+.1f  (wchars=%d ochars=%d)" % (
        p['i'], p['w_base'], p['o_base'], p['dy'], p['w_nchar'], p['o_nchar']))
dys = [p['dy'] for p in pairs]
if dys:
    import statistics
    print("\ndy: mean=%+.2f  spread(max-min)=%.1f" % (statistics.mean(dys), max(dys) - min(dys)))
    print("UNIFORM (small spread) => box-position/v-anchor offset; GROWING => line-height accum.")
print("wrote", OUT)
