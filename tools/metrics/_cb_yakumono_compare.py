# -*- coding: utf-8 -*-
"""Char-budget wall (c): compare per-line yakumono RENDER advances Word PDF vs Oxi dump.

Usage: python _cb_yakumono_compare.py <oxi_dump.json> <word_rt.pdf> [page]
Word advance = next_char.x0 - this_char.x0 (last char: x1-x0).
Oxi advance  = element 'w' (fragment width). Fragments are split at compression points.
Groups chars into lines by rounded y; aligns Word<->Oxi lines by text equality.
"""
import os, sys, json
sys.stdout.reconfigure(encoding='utf-8')
import fitz

YAK = set('、。，．「」『』（）〔〕【】《》〈〉｛｝［］・')
def is_yak(c): return c in YAK

OXI = os.path.abspath(sys.argv[1])
PDF = os.path.abspath(sys.argv[2])
ONLY_PAGE = next((int(a) for a in sys.argv[3:] if a.isdigit()), None)

# ---- Word PDF per-char ----
def word_lines(pdf):
    doc = fitz.open(pdf)
    pages = []
    for pno, page in enumerate(doc):
        d = page.get_text("rawdict")
        lines = []
        for blk in d.get("blocks", []):
            for ln in blk.get("lines", []):
                chars = []
                for sp in ln.get("spans", []):
                    for ch in sp.get("chars", []):
                        x0, y0, x1, y1 = ch["bbox"]
                        chars.append((ch["c"], x0, x1, y0, y1))
                if chars:
                    lines.append(chars)
        pages.append(lines)
    return pages

def line_adv(chars):
    """Return list of (char, advance) using x0 deltas."""
    out = []
    for i, (c, x0, x1, y0, y1) in enumerate(chars):
        if i + 1 < len(chars):
            adv = chars[i+1][1] - x0
        else:
            adv = x1 - x0
        out.append((c, adv))
    return out

# ---- Oxi dump per-char ----
def oxi_lines(dumpf):
    j = json.load(open(dumpf, encoding='utf-8'))
    pages = []
    for pg in j["pages"]:
        rows = {}
        for el in pg["elements"]:
            if el["type"] != "text": continue
            if not el["text"]: continue
            key = round(el["y"], 1)
            rows.setdefault(key, []).append(el)
        lines = []
        for y in sorted(rows):
            els = sorted(rows[y], key=lambda e: e["x"])
            chars = []
            for el in els:
                t = el["text"]
                w = el["w"]
                # multi-char fragment: split width evenly only for reporting non-yak;
                # yakumono are typically standalone fragments. Keep fragment-level for yak.
                if len(t) == 1:
                    chars.append((t, w))
                else:
                    per = w / len(t)
                    for c in t:
                        chars.append((c, per))
            lines.append(chars)
        pages.append(lines)
    return pages

wpages = word_lines(PDF)
opages = oxi_lines(OXI)
print("Word pages:", len(wpages), "Oxi pages:", len(opages))

def linetext(chars): return ''.join(c[0] for c in chars)

CLOSING = set('、。，．」』）〕】》〉｝］')
OPENING = set('「『（〔【《〈｛［')
def classify(idx, chars):
    """end | cluster | mid for the yak at idx in `chars` (list of (c,...))."""
    n = len(chars)
    c = chars[idx][0]
    if idx == n - 1: return 'end'
    prev_yak = idx > 0 and is_yak(chars[idx-1][0])
    next_yak = idx + 1 < n and is_yak(chars[idx+1][0])
    if prev_yak or next_yak: return 'cluster'
    return 'mid'

VERBOSE = '--verbose' in sys.argv
import collections
# Advance comparison EXCLUDES line-end yak (last-char ink width is not the advance).
agg = collections.defaultdict(lambda: [0, 0.0, 0])  # class -> [count, sum|d|, n_bad]
typ = collections.defaultdict(lambda: [0, 0.0, 0.0])  # chartype -> [count, sum_signed_d, sum|d|]
alln = [0, 0.0, 0]
# Per-line net packing: cumulative x0 position drift at the last char.
drift = []          # net |drift| at line end (start-aligned), char-identical lines
endyak_drift = []   # subset: lines ENDING in a yak
def ctype(c):
    if c in '、，': return '、'
    if c in '。．': return '。'
    if c in CLOSING: return 'close)'
    if c in OPENING: return 'open('
    if c == '・': return '・'
    return '?'
for pno in range(min(len(wpages), len(opages))):
    if ONLY_PAGE is not None and pno+1 != ONLY_PAGE: continue
    wl = wpages[pno]; ol = opages[pno]
    if VERBOSE: print(f"\n===== PAGE {pno+1}  (Word {len(wl)} lines, Oxi {len(ol)} lines) =====")
    used = set()
    for wi, wc in enumerate(wl):
        wt = linetext(wc)
        if not any(is_yak(c) for c in wt): continue
        match = None
        for oi, oc in enumerate(ol):
            if oi in used: continue
            if linetext(oc) == wt: match = oi; break
        if match is None:
            for oi, oc in enumerate(ol):
                if oi in used: continue
                ot = linetext(oc)
                if ot and (ot.startswith(wt[:8]) or wt.startswith(ot[:8])) and abs(len(ot)-len(wt))<=1:
                    match = oi; break
        if match is None: continue
        used.add(match)
        oc = ol[match]
        if len(oc) != len(wc): continue
        wadv = line_adv(wc)
        # cumulative position drift (start-aligned): pos_oxi[k]-pos_word[k]
        wx = 0.0; ox = 0.0; maxdrift = 0.0
        diffs = []
        for k in range(len(wc)):
            c = wc[k][0]
            wa = wadv[k][1]; oa = oc[k][1]
            is_last = (k == len(wc) - 1)
            if is_yak(c) and not is_last:
                dd = oa - wa
                cls = classify(k, wc)
                diffs.append((c, wa, oa, dd, cls))
                agg[cls][0] += 1; agg[cls][1] += abs(dd); agg[cls][2] += (abs(dd) > 0.6)
                t = ctype(c); typ[t][0]+=1; typ[t][1]+=dd; typ[t][2]+=abs(dd)
                alln[0] += 1; alln[1] += abs(dd); alln[2] += (abs(dd) > 0.6)
            if not is_last:
                wx += wa; ox += oa
                maxdrift = max(maxdrift, abs(ox - wx))
        drift.append(maxdrift)
        if is_yak(wc[-1][0]): endyak_drift.append(maxdrift)
        if '--drift-detail' in sys.argv and maxdrift > 2.5:
            # re-walk and print per-char position (start-aligned) Word vs Oxi
            wx2 = 0.0; ox2 = 0.0
            print(f"\n DRIFT {maxdrift:.2f} L{wi} n={len(wc)} end={wc[-1][0]!r}: {wt[:40]}")
            for k in range(len(wc) - 1):
                c = wc[k][0]
                d = ox2 - wx2
                mark = ' *' if abs(d) > 1.0 else ''
                if is_yak(c) or mark:
                    print(f"    [{k:2d}] {c} word_x={wx2:6.2f} oxi_x={ox2:6.2f} drift={d:+5.2f}{mark}")
                wx2 += wadv[k][1]; ox2 += oc[k][1]
        bad = [d for d in diffs if abs(d[3]) > 0.6]
        if VERBOSE and bad:
            print(f" L{wi} n={len(wc)} drift={maxdrift:.2f}: {wt[:42]}")
            for c, wa, oa, dd, cls in diffs:
                flag = "  <<<" if abs(dd) > 0.6 else ""
                print(f"     '{c}' [{cls:7}] word={wa:5.2f} oxi={oa:5.2f} d={dd:+5.2f}{flag}")

print("\n==== SUMMARY: yak ADVANCE error (char-identical lines, line-end excluded) ====")
print(f"{'class':9} {'n':>5} {'mean|d|':>8} {'n_bad(>0.6)':>12} {'%bad':>6}")
for cls in ('cluster', 'mid'):
    n, s, b = agg[cls]
    if n: print(f"{cls:9} {n:5d} {s/n:8.3f} {b:12d} {100*b/n:5.1f}%")
n, s, b = alln
if n: print(f"{'ALL':9} {n:5d} {s/n:8.3f} {b:12d} {100*b/n:5.1f}%")
print("\n  by char-type (signed mean = Oxi-Word; +=Oxi wider/under-compresses):")
print(f"  {'type':7} {'n':>5} {'mean_d':>8} {'mean|d|':>8}")
for t in ('、','。','close)','open(','・'):
    n,s,a = typ[t]
    if n: print(f"  {t:7} {n:5d} {s/n:+8.3f} {a/n:8.3f}")
import statistics as st
print("\n==== net per-line POSITION drift (cumulative, start-aligned) ====")
if drift:
    print(f"  all matched yak-lines : n={len(drift)} mean={st.mean(drift):.2f} max={max(drift):.2f} n(>2pt)={sum(d>2 for d in drift)}")
if endyak_drift:
    print(f"  lines ENDING in yak   : n={len(endyak_drift)} mean={st.mean(endyak_drift):.2f} max={max(endyak_drift):.2f} n(>2pt)={sum(d>2 for d in endyak_drift)}")
