# -*- coding: utf-8 -*-
"""Char-budget CELL derivation: from a _cb_gen_cell.py PDF (same 約物-rich BASE in a
single-cell table swept over cell WIDTH), extract per-line CELL break decisions and
約物 compression → derive Word's per-line CELL 約物 model (to port to Oxi's cell wrapper).

Mirrors _cb_derive.py (body) but avail = cell content width = cellw_pt - 2*cellMar_pt.
Cells segmented by the 甲 marker (BASE[0]). Cell N width = w0 + N*step (twips).

Usage: python _cb_derive_cell.py cb_cell.pdf [--sz 21] [--w0 6000] [--step 20]
       [--cellmar 108]
"""
import sys, os, math
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz
from collections import Counter

YAK = set('、。，．「」『』（）〔〕【】《》〈〉｛｝［］・')
CLOSING = set('、。，．」』）〕】》〉｝］・')   # right-aki 約物 (compressible by removing trailing aki)
OPENING = set('「『（〔【《〈｛［')
def is_yak(c): return c in YAK

a = sys.argv
def opt(name, d): return a[a.index(name)+1] if name in a else d
PDF = os.path.abspath(a[1])
sz = int(opt("--sz", "21"))
w0 = float(opt("--w0", "6000"))
step = float(opt("--step", "20"))
cellmar = float(opt("--cellmar", "108"))
fs = sz / 2.0
char_w = fs
cellmar_pt = cellmar / 20.0

doc = fitz.open(PDF)
# all lines across pages, reading order, with per-char (c,x0,x1)
lines = []
for page in doc:
    pls = []
    for blk in page.get_text("rawdict").get("blocks", []):
        for ln in blk.get("lines", []):
            chars = []
            for sp in ln.get("spans", []):
                for ch in sp.get("chars", []):
                    x0, y0, x1, y1 = ch["bbox"]
                    if ch["c"].strip():
                        chars.append((ch["c"], x0, x1, y0))
            if chars:
                pls.append(chars)
    pls.sort(key=lambda c: round(c[0][3], 1))
    lines.extend(pls)

# segment cells by 甲 (BASE[0])
START = '甲'
cells = []
cur = []
for chars in lines:
    if chars[0][0] == START:
        if cur: cells.append(cur)
        cur = [chars]
    elif cur:
        cur.append(chars)
if cur: cells.append(cur)
print(f"PDF: {len(lines)} lines, {len(cells)} cells (expected ~126)")

rows = []
adv_by = {}
for ci, clines in enumerate(cells):
    cellw_tw = w0 + ci * step
    avail = cellw_tw / 20.0 - 2 * cellmar_pt        # cell content width (pt)
    cap_nat = math.floor(avail / char_w + 1e-6)
    content_left = min(c[1] for c in clines[0])     # measured cell content-left
    content_right = content_left + avail
    for li, chars in enumerate(clines):
        is_last = (li == len(clines) - 1)
        seq = []
        for i, (c, x0, x1, y0) in enumerate(chars):
            adv = (chars[i+1][1] - x0) if i+1 < len(chars) else (x1 - x0)
            seq.append((c, round(adv, 3)))
        adv_by[(ci, li)] = seq
        n = len(chars)
        n_yak = sum(1 for c in chars if is_yak(c[0]))
        end_c = chars[-1][0]
        end_hang = chars[-1][2] - content_right
        rows.append(dict(ci=ci, li=li, last=is_last, avail=round(avail, 2),
                         n=n, cap=cap_nat, excess=n - cap_nat, n_yak=n_yak,
                         end=end_c, end_hang=round(end_hang, 2)))

full = [r for r in rows if not r['last']]
print(f"\nfull (justified) cell lines: {len(full)}")
print("excess (n_chars - natural_capacity) histogram:", dict(sorted(Counter(r['excess'] for r in full).items())))
oik = [r for r in full if r['excess'] > 0]
print(f"oikomi lines (excess>0, fit MORE than natural by 約物 compression): {len(oik)} / {len(full)}")
if oik:
    pair = Counter((r['excess'], r['n_yak']) for r in oik)
    print("  excess vs n_yak -> pt/約物 absorbed:")
    for (e, y), c in sorted(pair.items()):
        print(f"    excess={e} n_yak={y}: {c} lines  -> {e*char_w:.2f}pt absorbed, {(e*char_w/y) if y else 0:.2f}pt/約物")

# per-約物 advance distribution across ALL full-line 約物, by next-char class
solo_adv, pair_adv, open_adv = [], [], []
for r in full:
    seq = adv_by[(r['ci'], r['li'])]
    for i, (c, av) in enumerate(seq):
        if c in CLOSING and c != '・':
            nxt = seq[i+1][0] if i+1 < len(seq) else None
            if nxt and is_yak(nxt): pair_adv.append(av)
            else: solo_adv.append(av)
        elif c in OPENING:
            open_adv.append(av)
def stats(xs):
    if not xs: return "n=0"
    xs = sorted(xs); n = len(xs)
    return f"n={n} min={xs[0]:.2f} p10={xs[n//10]:.2f} median={xs[n//2]:.2f} mean={sum(xs)/n:.2f} max={xs[-1]:.2f}"
print(f"\n約物 advance (natural={char_w}):")
print(f"  closing/mark SOLO (next=non-約物): {stats(solo_adv)}")
print(f"  closing/mark before 約物 (PAIR):   {stats(pair_adv)}")
print(f"  OPENING bracket:                   {stats(open_adv)}")

# ぶら下げ: full lines ending in 約物 hanging past the cell content-right
hang = [r for r in full if is_yak(r['end']) and r['end_hang'] > 1.0]
print(f"\nぶら下げ (line-end 約物 hangs >1pt past cell margin): {len(hang)}/{len(full)}; "
      f"hang pt: {Counter(round(r['end_hang']) for r in hang).most_common(6)}")
print("full-line end-char (top):", Counter(r['end'] for r in full).most_common(8))

# oikomi vs oidashi DECISION: on lines that COULD oikomi (have 約物), did Word?
print(f"\n=== sample oikomi cell lines (約物 advances) ===")
for r in oik[:4]:
    seq = adv_by[(r['ci'], r['li'])]
    yak = [(c, av) for c, av in seq if is_yak(c)]
    print(f"  ci={r['ci']} li={r['li']} avail={r['avail']} n={r['n']} excess={r['excess']} n_yak={r['n_yak']} end={r['end']!r} hang={r['end_hang']}: 約物 {yak}")
