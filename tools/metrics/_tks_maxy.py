# -*- coding: utf-8 -*-
"""Per-page content line count + median line pitch, Word PDF vs Oxi dump.
Excludes header (<95) / footer (>770). Localizes per-line height deficit."""
import os, json, sys, tempfile, statistics
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz
tmp = tempfile.gettempdir()
PDF = os.path.join(tmp, 'tks_truth.pdf')
DUMP = os.path.join(tmp, 'tks_oxi_dump.json')
doc = fitz.open(PDF)
HMIN, HMAX = 95.0, 772.0

def word_line_ys(pg):
    ys = []
    for blk in pg.get_text('dict')['blocks']:
        if blk.get('type', 0) != 0:
            continue
        for ln in blk.get('lines', []):
            t = ''.join(s['text'] for s in ln['spans']).strip()
            if not t:
                continue
            y0 = min(s['bbox'][1] for s in ln['spans'])
            if y0 < HMIN or y0 > HMAX:
                continue
            ys.append(round(y0, 1))
    return sorted(set(ys))

d = json.load(open(DUMP, encoding='utf-8'))

def oxi_line_ys(pg):
    ys = set()
    for el in pg.get('elements', []):
        if el.get('type') == 'text' and HMIN <= el['y'] <= HMAX:
            ys.add(round(el['y'], 1))
    return sorted(ys)

def pitch(ys):
    if len(ys) < 2:
        return 0
    deltas = [b - a for a, b in zip(ys, ys[1:]) if 5 < (b - a) < 40]
    return statistics.median(deltas) if deltas else 0

lo = int(sys.argv[1]) if len(sys.argv) > 1 else 47
hi = int(sys.argv[2]) if len(sys.argv) > 2 else 64
print(f"{'pg':>3} {'Wlines':>6} {'Olines':>6} {'dLn':>4} {'Wpitch':>6} {'Opitch':>6} {'Wlast':>6} {'Olast':>6}")
twl = two = 0
for p in range(lo, hi):
    if p - 1 >= len(doc):
        break
    wys = word_line_ys(doc[p - 1])
    oys = oxi_line_ys(d['pages'][p - 1]) if p - 1 < len(d['pages']) else []
    twl += len(wys); two += len(oys)
    print(f"{p:>3} {len(wys):>6} {len(oys):>6} {len(oys)-len(wys):>+4} "
          f"{pitch(wys):>6.2f} {pitch(oys):>6.2f} {(wys[-1] if wys else 0):>6.1f} {(oys[-1] if oys else 0):>6.1f}")
print(f"TOTAL Word lines {twl}  Oxi lines {two}  (Oxi-Word {two-twl:+d})")
