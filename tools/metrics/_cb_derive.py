# -*- coding: utf-8 -*-
"""Char-budget derivation: from a _cb_gen.py PDF (same 約物-rich BASE swept over
right-indent), extract per-line break decisions and 約物 compression, then derive
Word's per-line oikomi budget.

Per FULL (non-last, justified) line:
  avail_pt   = page_content_width - right_indent(of this para)
  n_chars    = chars on the line
  cap_nat    = floor(avail_pt / char_w)         (natural fullwidth capacity)
  excess     = n_chars - cap_nat                (chars fit BEYOND natural = oikomi)
  n_yak      = 約物 count on the line
  end_char   = last char; end_hang = does it overflow content-right (ぶら下げ)
  The excess*char_w of width was absorbed by compressing the n_yak 約物.

Usage: python _cb_derive.py V1.pdf --sz 21 --pgw 11906 --marL 1418 --marR 1418
       (defaults match _cb_gen.py). char_w = sz/2/... no: sz is half-points,
       fontSize_pt = sz/2; fullwidth char_w = fontSize_pt.
"""
import sys, os, math
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz

YAK = set('、。，．「」『』（）〔〕【】《》〈〉｛｝［］・')
def is_yak(c): return c in YAK

def opt(a, name, dflt):
    return a[a.index(name)+1] if name in a else dflt

a = sys.argv
PDF = os.path.abspath(a[1])
sz = int(opt(a, "--sz", "21"))
pgw = int(opt(a, "--pgw", "11906"))
marL = int(opt(a, "--marL", "1418"))
marR = int(opt(a, "--marR", "1418"))
ind0 = float(opt(a, "--ind0", "0"))
step = float(opt(a, "--step", "0.5"))

fs = sz / 2.0                       # font size pt
char_w = fs                         # fullwidth natural advance (MS Mincho)
content_pt = (pgw - marL - marR) / 20.0   # page content width (pt), before right-indent

doc = fitz.open(PDF)
# gather all lines across pages in reading order: (chars=[(c,x0,x1)], y_origin)
lines = []
for page in doc:
    d = page.get_text("rawdict")
    page_lines = []
    for blk in d.get("blocks", []):
        for ln in blk.get("lines", []):
            chars = []
            for sp in ln.get("spans", []):
                for ch in sp.get("chars", []):
                    x0, y0, x1, y1 = ch["bbox"]
                    chars.append((ch["c"], x0, x1))
            if chars:
                page_lines.append((round(chars[0][1], 1), chars))
    page_lines.sort()  # by x0 of first char? no — by y. fix: sort by y
    # rebuild sorted by y (the tuple key is x0; redo)
    lines.extend(page_lines)

# Actually sort each page's lines by y (origin y). Re-extract with y.
lines = []
for page in doc:
    d = page.get_text("rawdict")
    pls = []
    for blk in d.get("blocks", []):
        for ln in blk.get("lines", []):
            chars = []
            for sp in ln.get("spans", []):
                for ch in sp.get("chars", []):
                    x0, y0, x1, y1 = ch["bbox"]
                    chars.append((ch["c"], x0, x1, y0))
            if chars:
                pls.append(chars)
    pls.sort(key=lambda c: round(c[0][3], 1))   # by first char y
    lines.extend(pls)

# Segment into paragraphs: a line whose first char == the doc's first char
# (unique to BASE position 0) starts a new paragraph.
START = lines[0][0][0] if lines else '甲'
paras = []
cur = []
for chars in lines:
    first = chars[0][0]
    if first == START:
        if cur:
            paras.append(cur)
        cur = [chars]
    else:
        if cur:
            cur.append(chars)
print(f"PDF: {len(lines)} lines, {len(paras)} paragraphs (expected ~ swept count)")
if cur:
    paras.append(cur)

# Map each paragraph to its right-indent (sweep order): para i -> ind0 + i*step pt
rows = []
for pi, plines in enumerate(paras):
    right_ind_pt = ind0 + pi * step
    avail = content_pt - right_ind_pt
    cap_nat = math.floor(avail / char_w + 1e-6)
    content_right = marL/20.0 + avail   # x of the right text boundary
    for li, chars in enumerate(plines):
        is_last = (li == len(plines) - 1)
        n = len(chars)
        n_yak = sum(1 for c in chars if is_yak(c[0]))
        end_c = chars[-1][0]
        end_x1 = chars[-1][2]
        end_hang = end_x1 - content_right   # >0 = hangs past margin (ぶら下げ)
        excess = n - cap_nat
        rows.append(dict(pi=pi, li=li, last=is_last, avail=round(avail,2),
                         n=n, cap=cap_nat, excess=excess, n_yak=n_yak,
                         end=end_c, end_hang=round(end_hang,2)))

# ---- per-line actual char advances (to see hang vs mid-compression) ----
# rebuild with advances keyed by (pi, li)
adv_by = {}
for pi, plines in enumerate(paras):
    for li, chars in enumerate(plines):
        seq = []
        for i, (c, x0, x1, y0) in enumerate(chars):
            adv = (chars[i+1][1] - x0) if i+1 < len(chars) else (x1 - x0)
            seq.append((c, round(adv, 2)))
        adv_by[(pi, li)] = seq

# Report: FULL lines (non-last) only — these are the justified break decisions.
full = [r for r in rows if not r['last']]
print(f"\nfull (justified) lines: {len(full)}")
from collections import Counter
exc = Counter(r['excess'] for r in full)
print("excess (n_chars - natural_capacity) histogram:", dict(sorted(exc.items())))
# excess>0 = oikomi (fit MORE than natural capacity by compressing 約物)
oik = [r for r in full if r['excess'] > 0]
print(f"oikomi lines (excess>0): {len(oik)} / {len(full)}")
if oik:
    # per-oikomi: width absorbed = excess*char_w ; per-約物 = absorbed/n_yak
    print("  excess vs n_yak (how many 約物 absorbed how many extra chars):")
    pair = Counter((r['excess'], r['n_yak']) for r in oik)
    for (e, y), c in sorted(pair.items()):
        per = (e*char_w/y) if y else 0
        print(f"    excess={e} n_yak={y}: {c} lines  -> {e*char_w:.2f}pt absorbed, {per:.2f}pt/約物")
# end-char of full lines (what char Word keeps at line end)
print("\nfull-line end-char (top):", Counter(r['end'] for r in full).most_common(8))
# ぶら下げ: full lines ending in 約物 that hang past margin
hang = [r for r in full if is_yak(r['end']) and r['end_hang'] > 1.0]
print(f"ぶら下げ (line-end 約物 hangs >1pt past margin): {len(hang)} lines; "
      f"end_hang values: {Counter(round(r['end_hang']) for r in hang).most_common(6)}")

# ---- mid-約物 advance on oikomi lines: compressed or natural? ----
print(f"\n=== 約物 advances on sample oikomi (excess=1) lines (char_w={char_w}) ===")
for r in oik[:4]:
    seq = adv_by[(r['pi'], r['li'])]
    yak_adv = [(c, av) for c, av in seq if is_yak(c)]
    print(f"  pi={r['pi']} li={r['li']} n={r['n']} n_yak={r['n_yak']} end={r['end']!r} hang={r['end_hang']}")
    print(f"     約物 advances: {yak_adv}")
    nonyak = [av for c, av in seq if not is_yak(c)]
    print(f"     non-約物 advance mean={sum(nonyak)/len(nonyak):.2f} (natural={char_w})")
# compare: excess=0 lines' 約物 advances (natural baseline)
print("=== 約物 advances on sample excess=0 (natural) lines ===")
for r in [x for x in full if x['excess']==0 and x['n_yak']>=2][:2]:
    seq = adv_by[(r['pi'], r['li'])]
    yak_adv = [(c, av) for c, av in seq if is_yak(c)]
    print(f"  pi={r['pi']} li={r['li']} n_yak={r['n_yak']} end={r['end']!r}: 約物 {yak_adv}")
