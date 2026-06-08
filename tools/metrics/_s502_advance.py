# -*- coding: utf-8 -*-
"""S502: measure Word's per-char advance for a given text run in a glyph json page, and
compare to the natural fullwidth advance (= font_size). Tells whether Word applies grid
char-spacing (advance>fs) or renders natural (advance==fs) for that line. cp932-safe."""
import json, io, sys

jp, pidx, needle = sys.argv[1], int(sys.argv[2]), sys.argv[3]
G = json.load(io.open(jp, encoding='utf-8'))['pages'][pidx]['glyphs']
chars = [g['char'] for g in G]
st = None
for i in range(len(chars) - len(needle) + 1):
    if ''.join(chars[i:i + len(needle)]) == needle:
        st = i; break
if st is None:
    print('NOT FOUND:', needle); sys.exit()
run = G[st:st + len(needle)]
fs = run[0].get('fs', run[0].get('font_size', 0))
print('text=%s  fs=%.1f  (natural fullwidth advance = fs)' % (needle, fs))
print('idx char     x      advance_from_prev')
prev = None
advs = []
for g in run:
    x = g['x']
    adv = (x - prev) if prev is not None else 0.0
    if prev is not None:
        advs.append(adv)
    print('    %s   %7.2f   %+.2f' % (g['char'], x, adv))
    prev = x
if advs:
    print('mean advance = %.3f   vs fs %.1f   -> grid_extra/char = %+.3f' % (
        sum(advs) / len(advs), fs, sum(advs) / len(advs) - fs))
