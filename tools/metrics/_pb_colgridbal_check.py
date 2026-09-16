# -*- coding: utf-8 -*-
"""Check candidate balance laws against the _pb_colgridbal arms (Word PDFs).

Law H (S1352 greedy on HEIGHTS, column-end space-after dropped, column-top
space-before kept):
    h = total / 2
    loop: fill column 1 greedily while the running height (with the last row's
          space-after dropped) stays <= h; if the rest fits h -> done; else h += pitch
Law S (grid SLOTS, ceil((h)/pitch - slack)).
"""
import os, sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_colgridbal_gen import ARMS, docx  # noqa: E402
import fitz  # noqa: E402

PITCH = 20.55


def word_split(label):
    page = fitz.open(docx(label)[:-5] + ".pdf")[0]
    mid = page.rect.width / 2
    n1 = 0
    for blk in page.get_text("dict")["blocks"]:
        for l in blk.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if not t or t in ("前のセクション", "次のセクション"):
                continue
            if l["bbox"][0] < mid:
                n1 += 1
    return n1


def rows_of(n, k, a, b):
    # (height without after, before, after) per row
    out = []
    for i in range(n):
        if i == k:
            out.append((PITCH, b / 20.0, a / 20.0))
        else:
            out.append((PITCH, 0.0, 0.0))
    return out


def law_h(rows):
    def height(i, last):
        h, bef, aft = rows[i]
        return h + bef + (0.0 if last else aft)
    total = sum(h + bef + aft for h, bef, aft in rows)
    hlim = total / 2.0
    for _ in range(len(rows) + 2):
        left = 0.0; split = 0
        for i in range(len(rows) - 1):
            cand = left + height(i, True)          # row i as the column's last row
            if cand > hlim + 0.001:
                break
            left = left + height(i, False) if i + 1 < len(rows) else cand
            split = i + 1
            left_as_last = cand
        if split == 0:
            split = 1; left_as_last = height(0, True)
        rest = sum(height(i, i == len(rows) - 1) for i in range(split, len(rows)))
        if rest <= hlim + 0.001:
            return split
        hlim += PITCH
    return len(rows) - 1


def law_s(rows, slack):
    import math
    slots = [max(1, math.ceil((h + bef + aft) / PITCH - slack)) for h, bef, aft in rows]
    total = sum(slots); left = 0; best = (10 ** 9, 1)
    for i in range(len(rows) - 1):
        left += slots[i]
        cost = abs(left - (total - left))
        if cost <= best[0]:
            best = (cost, i + 1)
    return best[1]


okh = oks1 = oks5 = 0
for label, n, k, a, b in ARMS:
    w = word_split(label)
    rows = rows_of(n, k, a, b)
    lh, ls1, ls5 = law_h(rows), law_s(rows, 0.1), law_s(rows, 0.5)
    okh += lh == w; oks1 += ls1 == w; oks5 += ls5 == w
    flag = "" if lh == w else "  <-- H differs"
    print("%-18s word=%d H=%d S0.1=%d S0.5=%d%s" % (label, w, lh, ls1, ls5, flag))
print("H %d/%d   S(0.1) %d/%d   S(0.5) %d/%d" % (okh, len(ARMS), oks1, len(ARMS), oks5, len(ARMS)))
