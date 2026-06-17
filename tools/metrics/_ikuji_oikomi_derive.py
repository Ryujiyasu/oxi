# -*- coding: utf-8 -*-
"""DERIVE Word's per-line jc=left oikomi/oidashi rule for ikujidetail (the S573
wall),正攻法: a clean labeled dataset from the Word PDF alone.

For each Word visual line (except a paragraph's last line), compute the EM-NATURAL
greedy break (fill from the line's first char at full-em widths until the content
right edge) using the SAME char stream Word laid out. Then compare to where Word
ACTUALLY broke:
  - Word breaks LATER than em-natural (kept char(s) em-natural would wrap) = OIKOMI
    (Word compressed 約物 / hung punctuation to fit more).
  - Word breaks == em-natural                                              = AGREE
  - Word breaks EARLIER than em-natural                                    = EARLY
para 152 「…交付す|る。」 is the canonical OIDASHI: Word wraps «る。» exactly where
em-natural wraps (AGREE). The lines S572 must fix are OIKOMI. The discriminator =
what separates OIKOMI from the AGREE-lines-that-still-overflow-by-a-char.

For each line we tabulate features of the FIRST char beyond Word's actual break
(the char Word pushed to the next line — the "oidashi char") and the LAST char(s)
Word kept (the potential "oikomi" it did): class, overflow-pt that em-natural sees,
trailing-約物 run, next-line first char, orphan size, #compressible 約物 on the line.
"""
import os, sys, tempfile, unicodedata
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz

PDF = os.path.join(tempfile.gettempdir(), 'ikd_truth.pdf')
FS = 11.0
EDGE = 538.6
CLOSERS = '」』）】〕》〉｝］、。，．'
OPENERS = '「『（【〔《〈｛［'
YAK = CLOSERS + OPENERS + '・：；'
COMPRESSIBLE = '、。，．）」』】〕》〉｝］'   # closers + commas/periods (line-end hangable)


def em(c):
    if c in ' \t': return FS / 2.0
    if c == '　': return FS
    if '0' <= c <= '9' or 'A' <= c <= 'Z' or 'a' <= c <= 'z': return FS / 2.0
    return FS if unicodedata.east_asian_width(c) in ('W', 'F', 'A') else FS / 2.0


def cls(c):
    if c in '、，': return 'comma'
    if c in '。．': return 'period'
    if c == '・': return 'naka'
    if c in OPENERS: return 'opener'
    if c in CLOSERS.replace('、。，．', ''): return 'closer'
    if c and ('0' <= c <= '9' or 'a' <= c <= 'z' or 'A' <= c <= 'Z'
              or '０' <= c <= '９'): return 'latin'
    return 'cjk'


# Build Word visual lines (Y-cluster), each = list of (x0,x1,char) in reading order.
doc = fitz.open(PDF)
vlines = []
for pno in range(len(doc)):
    rd = doc[pno].get_text('rawdict')
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
    for k in sorted(rows):
        r = sorted(rows[k])
        vlines.append({'page': pno+1, 'x0': r[0][0],
                       'chars': [c for _, _, c in r],
                       'xs': [(x0, x1) for x0, x1, _ in r]})

# A paragraph break = the next line's text doesn't continue the same paragraph.
# Heuristic: a line is a paragraph's LAST line if the next visual line starts a
# new para. We can't see para marks in PDF, but for the oikomi test we only need:
# for line i, the chars that em-natural would pull from line i+1. We treat line
# i+1 as the continuation candidate (works for wrapped paras; a false continuation
# at a real para boundary just yields AGREE/early which we can filter by big x gap).
rows_out = []
for i in range(len(vlines) - 1):
    L = vlines[i]; N = vlines[i+1]
    if N['page'] != L['page']:
        continue
    # only consider lines that look "full" (Word filled near the edge): else it's
    # a paragraph's last (short) line — not a wrap decision.
    word_end_x = L['xs'][-1][1]
    if word_end_x < EDGE - 2 * FS:
        continue                       # short line = para end, skip
    avail = EDGE - L['x0']
    # em-natural greedy break over L.chars + N.chars
    seq = L['chars'] + N['chars']
    cum = 0.0
    nat_break = len(seq)               # index where natural wraps (char goes to next line)
    for j, c in enumerate(seq):
        cum += em(c)
        if cum > avail + 0.5:
            nat_break = j
            break
    word_break = len(L['chars'])        # Word's actual break index in seq
    # the char Word PUSHED to the next line (first char of N)
    oidashi_char = N['chars'][0] if N['chars'] else ''
    # the char(s) Word KEPT past em-natural (if word_break > nat_break = OIKOMI)
    kept = seq[nat_break:word_break]
    # em-natural's pushed char (what em-natural would wrap)
    nat_pushed = seq[nat_break] if nat_break < len(seq) else ''
    # overflow that em-natural sees on the kept run (how far past avail)
    overflow = (sum(em(c) for c in seq[:word_break]) - avail) if word_break <= len(seq) else 0.0
    label = ('OIKOMI' if word_break > nat_break
             else ('EARLY' if word_break < nat_break else 'AGREE'))
    ncomp = sum(1 for c in L['chars'] if c in COMPRESSIBLE)
    # trailing compressible run at Word's line end
    trail = 0
    for c in reversed(L['chars']):
        if c in COMPRESSIBLE: trail += 1
        else: break
    rows_out.append({
        'page': L['page'], 'label': label,
        'kept': ''.join(kept), 'nat_pushed': nat_pushed, 'oidashi_char': oidashi_char,
        'kept_cls': cls(kept[0]) if kept else '', 'nat_pushed_cls': cls(nat_pushed),
        'overflow': overflow, 'ncomp': ncomp, 'trail_comp': trail,
        'line_end': L['chars'][-1], 'line_end_cls': cls(L['chars'][-1]),
        'tail': ''.join(L['chars'][-8:]),
    })

from collections import Counter
oik = [r for r in rows_out if r['label'] == 'OIKOMI']
agr = [r for r in rows_out if r['label'] == 'AGREE']
ear = [r for r in rows_out if r['label'] == 'EARLY']
print(f"full wrap-lines: {len(rows_out)}  OIKOMI={len(oik)}  AGREE={len(agr)}  EARLY={len(ear)}")

print("\n=== OIKOMI: how many chars Word kept past em-natural ===")
print("  kept-length:", dict(Counter(len(r['kept']) for r in oik)))
print("  kept[0] class:", dict(Counter(r['kept_cls'] for r in oik)))
print("  overflow pt (em-natural sees): ", end='')
import statistics
ov = [r['overflow'] for r in oik]
if ov: print(f"median={statistics.median(ov):.1f} p90={sorted(ov)[int(len(ov)*0.9)]:.1f} max={max(ov):.1f}")

print("\n=== AGREE (oidashi-or-exact): the char Word pushed (= em-natural pushed) ===")
print("  pushed class:", dict(Counter(r['nat_pushed_cls'] for r in agr)))
# AGREE lines where the pushed char is REGULAR and overflow is ~1 char = the para-152 family
agr_reg = [r for r in agr if r['nat_pushed_cls'] in ('cjk', 'latin')]
print(f"  AGREE with REGULAR pushed char (para-152 family): {len(agr_reg)}/{len(agr)}")

print("\n=== KEY CONTRAST: OIKOMI kept-char class vs AGREE pushed-char class ===")
print("  When Word KEEPS past natural (oikomi), the kept char is:", dict(Counter(r['kept_cls'] for r in oik)))
print("  When Word PUSHES at natural (agree),  the pushed char is:", dict(Counter(r['nat_pushed_cls'] for r in agr)))

print("\n--- sample OIKOMI lines (kept='what Word kept past natural') ---")
for r in oik[:18]:
    print(f"  p{r['page']} kept='{r['kept']}' ({r['kept_cls']}) ov={r['overflow']:.1f} ncomp={r['ncomp']} trail={r['trail_comp']} …{r['tail']}")
print("\n--- sample AGREE lines with REGULAR pushed char (para-152 type) ---")
for r in agr_reg[:18]:
    print(f"  p{r['page']} pushed='{r['nat_pushed']}' ({r['nat_pushed_cls']}) ov={r['overflow']:.1f} ncomp={r['ncomp']} trail={r['trail_comp']} …{r['tail']}|{r['oidashi_char']}")
