# -*- coding: utf-8 -*-
"""Per-LINE break-divergence finder: align Oxi-dump lines with Word-PDF lines
(char-stream) and report every line where Oxi's break point differs from Word
(over-fit = Oxi fits a char Word wraps; under-fit = Oxi wraps a char Word fits).

The per-line precision foundation for the demand-aware breaker. Run with the
Oxi dump at C:/tmp/n_def.json (render with the desired env, --dump-layout).

Usage: python _nedo_linebreak.py [first|all]   (default: first divergence + context)
"""
import json, sys, fitz, difflib
from collections import defaultdict
sys.stdout.reconfigure(encoding="utf-8")
PDF = r"C:\tmp\nedocontract_word.pdf"
OXI = r"C:/tmp/n_def.json"
HEADER_SUBSTR = "一般再委託用"
YAK = set("（）「」『』〔〕【】、。，．・")
MODE = sys.argv[1] if len(sys.argv) > 1 else "first"

def word_lines():
    doc = fitz.open(PDF); out = []
    for pi in range(doc.page_count):
        d = doc.load_page(pi).get_text("rawdict")
        for blk in d["blocks"]:
            if blk.get("type") != 0: continue
            for ln in blk.get("lines", []):
                cs = "".join(c["c"] for sp in ln.get("spans", []) for c in sp.get("chars", []))
                cs = cs.rstrip()
                if cs and HEADER_SUBSTR not in cs:
                    out.append((pi+1, cs))
    return out

def oxi_lines():
    d = json.load(open(OXI, encoding="utf-8"))
    out = []
    for pg in d["pages"]:
        lines = defaultdict(list)
        for e in pg["elements"]:
            if e.get("type") == "text" and e.get("text", "").strip():
                lines[round(e["y"])].append(e)
        for y in sorted(lines):
            txt = "".join(e["text"] for e in sorted(lines[y], key=lambda e: e["x"])).rstrip()
            if txt and HEADER_SUBSTR not in txt:
                out.append((pg["page"], txt))
    return out

wl = word_lines(); ol = oxi_lines()
# build char streams with per-char (which-line-index) tags
wch = []; wtag = []
for i,(pg,t) in enumerate(wl):
    for c in t: wch.append(c); wtag.append(i)
och = []; otag = []
for i,(pg,t) in enumerate(ol):
    for c in t: och.append(c); otag.append(i)

# align
sm = difflib.SequenceMatcher(None, och, wch, autojunk=False)
# map: for each matched char, (oxi_line_idx, word_line_idx)
# A break divergence at a Word line boundary: the char AFTER a word-line-end
# is on a NEW word line; find where Oxi puts it (same line = match, diff = divergence)
pairs = []  # (oxi_line, word_line) for matched chars in order
for a, b, size in sm.get_matching_blocks():
    for k in range(size):
        pairs.append((otag[a+k], wtag[b+k]))

# For each Word line, the set of Oxi lines its chars fall on. A break divergence:
# the LAST char of word line W and the FIRST char of word line W+1 are on the
# SAME oxi line (Oxi over-fit: didn't break where Word did) OR Oxi broke earlier.
# Walk word-line transitions.
div = []
# group matched pairs by consecutive runs; detect at each word-line transition
prev_w = None; prev_o = None
for o, w in pairs:
    if prev_w is not None and w != prev_w:
        # transition from word line prev_w to w
        if o == prev_o:
            # the new word line's first matched char is on the SAME oxi line as
            # the prev word line's last char => Oxi did NOT break here = OVER-FIT
            div.append(("OVERFIT", prev_w, prev_o))
        # (under-fit = Oxi broke but Word didn't, harder to detect here)
    prev_w, prev_o = w, o

# also detect under-fit: one Word line maps to >1 oxi line
w2o = defaultdict(set)
for o, w in pairs: w2o[w].add(o)
under = [(w, sorted(os)) for w, os in w2o.items() if len(os) > 1]

print(f"Word lines: {len(wl)}  Oxi lines: {len(ol)}")
print(f"\n=== OVER-FIT divergences (Oxi fits past Word's break) ===")
shown = 0
for kind, w, o in div:
    wpg, wtxt = wl[w]; otxt = ol[o] if o < len(ol) else "?"
    nxt = wl[w+1][1][:1] if w+1 < len(wl) else ""
    yk = [c for c in wtxt if c in YAK]
    print(f"  Wp{wpg} Wline{w} end='{wtxt[-3:]}' nextWchar='{nxt}' 約物={yk[-4:]}")
    print(f"        Oxi line: ...{otxt[-12:]}")
    shown += 1
    if MODE == "first" and shown >= 3: break
print(f"\n=== UNDER-FIT (1 Word line -> multiple Oxi lines; Oxi wraps early) ===")
for w, os in under[:6 if MODE=="first" else 999]:
    print(f"  Wline{w} p{wl[w][0]} -> oxi lines {os}: {wl[w][1][:36]}")
print(f"\ntotals: overfit={len(div)} underfit={len(under)}")
