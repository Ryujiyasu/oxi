# -*- coding: utf-8 -*-
"""Label EVERY nedocontract Word-PDF line oikomi (compressed-to-fit) vs natural,
with features, to DERIVE the demand-aware break's oikomi/oidashi discriminator.

For each full content line: per-約物 compression (12 - advance), total line
compression, the trailing char + next-line first char (the wrap boundary),
n_period, n_yakumono, char count. Word's break = GREEDY-at-natural + demand
compression on SOME lines. The question: what feature makes Word compress a
line (oikomi, fit the next char) vs break it natural (oidashi, wrap)?
"""
import sys, fitz, collections
sys.stdout.reconfigure(encoding="utf-8")
PDF = r"C:\tmp\nedocontract_word.pdf"
YAK = set("（）「」『』〔〕【】《》〈〉｛｝［］、。，．・")
PERIOD = set("。．")
CLOSING = set("）」』〕】》〉｝］")
OPENING = set("（「『〔【《〈｛［")

def page_lines(page):
    d = page.get_text("rawdict"); out = []
    for blk in d["blocks"]:
        if blk.get("type") != 0: continue
        for ln in blk.get("lines", []):
            chars = []
            for sp in ln.get("spans", []):
                for c in sp.get("chars", []):
                    bb = c["bbox"]; chars.append((c["c"], bb[0], bb[2]))
            if chars: out.append(chars)
    return out

doc = fitz.open(PDF)
# collect all lines across pages in reading order
all_lines = []
for pi in range(len(doc)):
    for chars in page_lines(doc[pi]):
        all_lines.append((pi+1, chars))

# fullwidth ~12.0; line is "full content" if it spans most of the text width
rows = []
for i,(pg,chars) in enumerate(all_lines):
    # advances: x1-x0 per char (last char = ink, skip for adv)
    text = "".join(c[0] for c in chars).rstrip()
    if len(text) < 20: continue  # skip short/heading lines
    x0 = chars[0][1]; x1 = chars[-1][2]
    width = x1 - x0
    if width < 360: continue  # only near-full lines
    # per-約物 compression (use next char's x0 - this char's x0 = advance)
    comp_total = 0.0; yak_comps = []
    for j in range(len(chars)-1):
        ch = chars[j][0]
        if ch in YAK:
            adv = chars[j+1][1] - chars[j][1]
            comp = 12.0 - adv
            if comp > 0.3:
                comp_total += comp; yak_comps.append((ch, round(comp,2)))
    nper = sum(1 for c in text if c in PERIOD)
    nyak = sum(1 for c in text if c in YAK)
    # next-line first non-space char
    nxt = ""
    if i+1 < len(all_lines):
        nt = "".join(c[0] for c in all_lines[i+1][1]).lstrip()
        nxt = nt[0] if nt else ""
    rows.append(dict(pg=pg, n=len(text), comp=round(comp_total,2), nper=nper,
                     nyak=nyak, end=text[-2:], nxt=nxt, yak=yak_comps[:4]))

# summary
comp_lines = [r for r in rows if r['comp'] > 1.0]
nat_lines  = [r for r in rows if r['comp'] <= 1.0]
print(f"total near-full lines: {len(rows)}")
print(f"  OIKOMI (comp>1.0): {len(comp_lines)} ({100*len(comp_lines)//max(1,len(rows))}%)")
print(f"  NATURAL (comp<=1.0): {len(nat_lines)}")
print(f"\nOIKOMI lines comp distribution:")
import statistics
cs = sorted(r['comp'] for r in comp_lines)
if cs:
    print(f"  min {cs[0]} median {statistics.median(cs):.2f} p90 {cs[int(len(cs)*0.9)]:.2f} max {cs[-1]}")
print(f"\noikomi lines n_period: {collections.Counter(r['nper'] for r in comp_lines)}")
print(f"natural lines n_period: {collections.Counter(r['nper'] for r in nat_lines)}")
print(f"\noikomi: next char is closing-bracket? {sum(1 for r in comp_lines if r['nxt'] in CLOSING)}/{len(comp_lines)}")
print(f"oikomi: next char is opening-bracket? {sum(1 for r in comp_lines if r['nxt'] in OPENING)}/{len(comp_lines)}")
print(f"oikomi: next char is period/comma? {sum(1 for r in comp_lines if r['nxt'] in PERIOD or r['nxt']=='、')}/{len(comp_lines)}")
print(f"\n=== sample OIKOMI lines (Word compressed to fit) ===")
for r in comp_lines[:14]:
    print(f"  pg{r['pg']} n={r['n']} comp={r['comp']} nper={r['nper']} end='{r['end']}' nxt='{r['nxt']}' yak={r['yak']}")
