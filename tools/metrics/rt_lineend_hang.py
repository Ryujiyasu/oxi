# -*- coding: utf-8 -*-
"""For each content-matched line, tabulate the line-END char and the Word-vs-Oxi line-end-x
difference (calibrated). Tests the hanging-punctuation (burasagari) hypothesis: lines ending in
closing punctuation where Word hangs past margin but Oxi does not -> large +diff. cp932-safe."""
import json, sys, collections
d = json.load(open(sys.argv[1], encoding='utf-8'))
W, O = d['word'], d['oxi']
cal = float(sys.argv[2]) if len(sys.argv) > 2 else 52.3  # word->oxi x offset (median)
n = min(len(W), len(O))
by_end = collections.defaultdict(list)
rows = []
for i in range(n):
    w, o = W[i], O[i]
    if not o['text'] or not w['text']:
        continue
    oend = o['text'][-1]
    wend = w['text'][-1]
    if oend != wend:
        continue  # only matched-end lines
    # last-char x in common (oxi) frame
    w_x1 = w['x1'] + cal
    o_x1 = o['x1']
    diff = w_x1 - o_x1  # >0 => Word's last char is RIGHT of Oxi's (Word hangs more)
    by_end[oend].append(diff)
    rows.append({'i': i, 'end': 'U+%04X' % ord(oend), 'wx1': round(w_x1, 1), 'ox1': round(o_x1, 1),
                 'diff': round(diff, 1), 'n': o['n']})

print("per-line: Word_lastx - Oxi_lastx (calib). >0 = Word hangs further right than Oxi")
for r in rows:
    flag = '  <== WORD HANGS, OXI DOESNT' if r['diff'] >= 4 else ('  (both hang ~equal)' if r['diff'] <= 1.5 and r['ox1'] > 535 else '')
    print("  L%2d end=%s wx1=%6.1f ox1=%6.1f diff=%+5.1f n=%d%s" % (
        r['i'], r['end'], r['wx1'], r['ox1'], r['diff'], r['n'], flag))

print("\n=== grouped by line-end char (diff = Word hang beyond Oxi) ===")
for ch, ds in sorted(by_end.items(), key=lambda kv: -sum(kv[1]) / len(kv[1])):
    import statistics
    print("  U+%04X count=%2d  diff mean=%+5.2f  vals=%s" % (
        ord(ch), len(ds), statistics.mean(ds), [round(x, 1) for x in ds][:12]))
