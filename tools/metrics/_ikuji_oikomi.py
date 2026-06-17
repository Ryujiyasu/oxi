# -*- coding: utf-8 -*-
"""Statistical derivation of Word's per-line oikomi (compress 約物 to fit) vs
oidashi (wrap) decision for ikujidetail — the S573 wall.

For each Word-PDF body line: natural width (Σ em, 約物 at full em) vs actual
rendered width. absorption = natural − actual. A "full" line reaches the content
right edge (≈538.6). Classify full lines: COMPRESSED (absorb > 1pt = oikomi) vs
NATURAL (absorb ≈ 0). Then tabulate features (line-end char class, next-line
first-char class, #約物, overflow-if-not-compressed) to find what separates the
two — i.e. when Word chooses oikomi vs oidashi.

The decisive contrast: para 152 «…交付す|る。» is OIDASHI (Word wraps «る。»)
while many lines are OIKOMI. Both have compressible 約物 — the discriminator is
the unsolved lever.
"""
import os, sys, tempfile, unicodedata
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz

PDF = os.path.join(tempfile.gettempdir(), 'ikd_truth.pdf')
FS = 11.0
EDGE = 538.6   # content right edge (pgW 595.3 − rightMargin 56.7)
CLOSERS = '」』）】〕》〉｝］、。，．'
OPENERS = '「『（【〔《〈｛［'
YAK = CLOSERS + OPENERS + '・'


def cls(c):
    if c in '、，': return 'comma'
    if c in '。．': return 'period'
    if c == '・': return 'naka'
    if c in OPENERS: return 'opener'
    if c in CLOSERS.replace('、。，．', ''): return 'closer'
    if c and ('0' <= c <= '9' or 'a' <= c <= 'z' or 'A' <= c <= 'Z'
              or '０' <= c <= '９'): return 'latin'
    return 'cjk'


def is_fw(c):
    return unicodedata.east_asian_width(c) in ('W', 'F', 'A')


doc = fitz.open(PDF)
lines = []   # per line: dict
for pno in range(len(doc)):
    pg = doc[pno]
    rd = pg.get_text('rawdict')
    chars = []
    for blk in rd['blocks']:
        if blk.get('type', 0) != 0:
            continue
        for ln in blk.get('lines', []):
            for sp in ln['spans']:
                for ch in sp['chars']:
                    b = ch['bbox']
                    if b[1] < 65 or b[1] > 795:
                        continue
                    chars.append((round((b[1]+b[3])/2, 1), b[0], b[2], ch['c']))
    chars.sort(key=lambda t: (round(t[0]/3), t[1]))
    rows = {}
    for yc, x0, x1, c in chars:
        rows.setdefault(round(yc/3), []).append((x0, x1, c))
    ks = sorted(rows)
    for i, k in enumerate(ks):
        r = sorted(rows[k])
        txt = ''.join(c for _, _, c in r)
        if len(r) < 3:
            continue
        x0 = r[0][0]; x1 = r[-1][1]
        actual = x1 - x0
        nat = sum(FS if is_fw(c) else FS/2.0 for _, _, c in r)
        full = x1 >= EDGE - FS   # reaches the right edge
        # the next visual line's first char (for kinsoku context)
        lines.append({'page': pno+1, 'txt': txt, 'x0': x0, 'x1': x1,
                      'actual': actual, 'nat': nat, 'absorb': nat - actual,
                      'full': full, 'end': txt[-1], 'nyak': sum(1 for c in txt if c in YAK)})

full = [l for l in lines if l['full']]
comp = [l for l in full if l['absorb'] > 1.0]    # oikomi: Word compressed 約物
nat = [l for l in full if l['absorb'] <= 1.0]     # oidashi/natural
print(f"body lines: {len(lines)}   full lines: {len(full)}   "
      f"OIKOMI(absorb>1): {len(comp)}   NATURAL(absorb<=1): {len(nat)}")
import statistics
if comp:
    print(f"  oikomi absorb: median={statistics.median(l['absorb'] for l in comp):.2f} "
          f"max={max(l['absorb'] for l in comp):.2f}")

from collections import Counter
print("\n--- OIKOMI lines: last-char class (what Word KEEPS by compressing) ---")
for k, n in Counter(cls(l['end']) for l in comp).most_common():
    print(f"   {k}: {n}")
print("--- NATURAL lines: last-char class ---")
for k, n in Counter(cls(l['end']) for l in nat).most_common():
    print(f"   {k}: {n}")

# The decisive question: do OIKOMI lines end in 約物 (Word compresses to keep a
# trailing 約物) while OIDASHI lines would need to keep a REGULAR char?
print("\n--- OIKOMI lines ending in 約物 vs regular ---")
oik_yak = sum(1 for l in comp if l['end'] in YAK)
print(f"   end in 約物: {oik_yak}/{len(comp)}   end regular: {len(comp)-oik_yak}/{len(comp)}")

print("\n--- sample OIKOMI lines (absorb, end, text tail) ---")
for l in sorted(comp, key=lambda l: -l['absorb'])[:12]:
    print(f"   p{l['page']} absorb={l['absorb']:.1f} end='{l['end']}' …{l['txt'][-18:]}")
print("\n--- sample NATURAL full lines (absorb≈0) ---")
for l in nat[:8]:
    print(f"   p{l['page']} absorb={l['absorb']:.1f} end='{l['end']}' …{l['txt'][-18:]}")
