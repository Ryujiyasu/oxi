# -*- coding: utf-8 -*-
"""Per-element localize: for the glyphs the tblInd fix SHIFTS, decide per glyph whether
Word matches the LEGACY x (literal, no absorption) or the FIX x (absorption). Groups by the
shift magnitude so we can tell WHICH tables in a doc absorb vs not. Inputs: legacy Oxi dump,
fix Oxi dump (same glyph sequence, x differs), Word PDF glyphs. cp932-safe, ASCII out."""
import json, sys, difflib, statistics
from collections import defaultdict

leg = json.load(open(sys.argv[1], encoding='utf-8'))
fix = json.load(open(sys.argv[2], encoding='utf-8'))
wrd = json.load(open(sys.argv[3], encoding='utf-8'))
page = int(sys.argv[4]) if len(sys.argv) > 4 else 0

lg = [g for g in leg['pages'][page]['glyphs'] if g['char'].strip()]
fg = [g for g in fix['pages'][page]['glyphs'] if g['char'].strip()]
wg = [g for g in wrd['pages'][page]['glyphs'] if g['char'].strip()]
if len(lg) != len(fg):
    print('legacy/fix glyph count differ', len(lg), len(fg)); sys.exit()

# pair legacy<->fix by index (same sequence); legacy x + fix x per glyph
pairs = [(a['char'], a['x'], b['x']) for a, b in zip(lg, fg)]
# match Word to legacy by content
sm = difflib.SequenceMatcher(None, [g['char'] for g in lg], [g['char'] for g in wg], autojunk=False)
matched = []
for tag, i1, i2, j1, j2 in sm.get_opcodes():
    if tag == 'equal':
        for d in range(i2 - i1):
            ch, lx, fx = pairs[i1 + d]
            wx = wg[j1 + d]['x']
            matched.append((ch, lx, fx, wx))
# global cal: median (legacy_x - word_x) over UNSHIFTED glyphs (lx==fx)
uns = [lx - wx for ch, lx, fx, wx in matched if abs(lx - fx) < 0.05]
cal = statistics.median(uns) if uns else 0.0
print('matched %d  unshifted %d  cal(legacy-word)=%.2f' % (len(matched), len(uns), cal))

# for shifted glyphs, bucket by shift magnitude; decide legacy vs fix closer to word
buckets = defaultdict(lambda: {'n': 0, 'leg_closer': 0, 'fix_closer': 0, 'wmlx': [], 'wmfx': []})
for ch, lx, fx, wx in matched:
    shift = round(fx - lx, 1)
    if abs(shift) < 0.05:
        continue
    wxc = wx + cal  # word in legacy frame
    dl = abs(lx - wxc); df = abs(fx - wxc)
    b = buckets[shift]
    b['n'] += 1
    if dl < df: b['leg_closer'] += 1
    else: b['fix_closer'] += 1
    b['wmlx'].append(lx - wxc); b['wmfx'].append(fx - wxc)
print('\nshift  n    legacy_closer fix_closer  | mean(leg-word) mean(fix-word)  verdict')
for shift in sorted(buckets):
    b = buckets[shift]
    ml = statistics.mean(b['wmlx']); mf = statistics.mean(b['wmfx'])
    verdict = 'ABSORB (fix right)' if b['fix_closer'] > b['leg_closer'] else 'LITERAL (legacy right)'
    print('%+5.1f %4d   %5d        %5d      | %+.2f         %+.2f        %s'
          % (shift, b['n'], b['leg_closer'], b['fix_closer'], ml, mf, verdict))
