# -*- coding: utf-8 -*-
"""Oikomi census on a Word PDF: which full lines were pulled in past the cell
count, by how much per mark, and which lines were NOT pulled in although marks
were available.  Derives the pull-in cap from the document itself.

Usage: python tools/metrics/_oikomi_census.py <word.pdf> [--pages=a,b] [--show]

A line's pitch is the median advance between its full-width characters; its
column width is the 90th-percentile width of lines of its class (two-column
lines narrower than 250pt, single-column lines wider).  cells = width / pitch.
need = full chars + 0.5 per half-width char - 0.5 per structural yakumono pair
(a closing/mid mark followed by another mark collapses one aki).  A GRANT is a
line whose need exceeds the cells; a REFUSE is a line whose need fits but which
would exceed the cells by pulling in the first character of the next line (a
normal character -- a kinsoku mark is not a choice).  per-mark = the demand
divided by the number of marks on the line (each mark has 0.5em of aki; a
middle dot has 0.25 on each side).
"""
import sys
from collections import Counter

import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
pdf = sys.argv[1]
pages = None
show = "--show" in sys.argv
colw_arg = None
colw1_arg = None
for a in sys.argv[2:]:
    if a.startswith("--pages="):
        pages = [int(x) for x in a.split("=")[1].split(",")]
    if a.startswith("--colw="):
        colw_arg = float(a.split("=")[1])
    if a.startswith("--colw1="):
        colw1_arg = float(a.split("=")[1])

OPEN = set("（「『【［〔《〈")
CLOSE = set("）」』】］〕》〉")
MID = set("、。，．・：；")


def is_full(c):
    o = ord(c)
    return o >= 0x2E80 or c == "　" or 0x2460 <= o <= 0x24FF or 0x2160 <= o <= 0x217F


def cells_of(text):
    n = 0.0
    pairs = 0
    marks = []
    prev = None
    for c in text:
        if c == " ":
            continue
        n += 1.0 if is_full(c) else 0.5
        if c in MID or c in OPEN or c in CLOSE:
            marks.append(c)
        if prev is not None and (prev in CLOSE or prev in MID) and (c in OPEN or c in CLOSE or c in MID):
            pairs += 1
        prev = c
    return n, pairs, marks


def median(v):
    v = sorted(v)
    return v[len(v) // 2] if v else 0.0


def main():
    doc = fitz.open(pdf)
    plist = pages or list(range(1, len(doc) + 1))
    lines = []
    for pno in plist:
        page = doc[pno - 1]
        for b in page.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                chars = [c for s in l["spans"] for c in s["chars"]]
                t = "".join(c["c"] for c in chars)
                if not t.strip() or len(chars) < 2:
                    continue
                advs = [chars[k + 1]["origin"][0] - chars[k]["origin"][0]
                        for k in range(len(chars) - 1)
                        if is_full(chars[k]["c"]) and is_full(chars[k + 1]["c"])]
                pitch = median(advs)
                if pitch <= 0:
                    continue
                lines.append((pno, round(l["bbox"][0], 1), round(l["bbox"][1], 1), round(l["bbox"][2], 1), t, pitch))
    # a line Word squeezed reads a low pitch; take the max over the line and its neighbours
    fixed = []
    for i, rec in enumerate(lines):
        nb = [lines[j][5] for j in (i - 1, i + 1) if 0 <= j < len(lines) and lines[j][0] == rec[0]]
        fixed.append(rec[:5] + (max([rec[5]] + nb),))
    lines = fixed
    jc_of = {}
    docx_arg = next((a.split("=", 1)[1] for a in sys.argv if a.startswith("--docx=")), None)
    if docx_arg:
        import re
        import zipfile
        xml = zipfile.ZipFile(docx_arg).read("word/document.xml").decode("utf-8")
        for pm in re.finditer(r"<w:p[ >].*?</w:p>", xml, re.S):
            body = pm.group(0)
            ppr = re.search(r"<w:pPr>.*?</w:pPr>", body, re.S)
            m = re.search(r'<w:jc w:val="(\w+)"', ppr.group(0)) if ppr else None
            jc = m.group(1) if m else "-"
            text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", body)).replace(" ", "").replace("　", "")
            for k in range(0, max(1, len(text) - 7)):
                jc_of.setdefault(text[k:k + 8], jc)
    xs = Counter(x0 for _, x0, _, _, _, _ in lines)
    col_x = [x for x, _ in xs.most_common(2)]
    w2 = sorted(x1 - x0 for _, x0, _, x1, _, _ in lines if x0 in col_x and x1 - x0 < 250)
    w1 = sorted(x1 - x0 for _, x0, _, x1, _, _ in lines if x0 in col_x and x1 - x0 >= 250)
    colw2 = colw_arg or (w2[int(len(w2) * 0.9)] if w2 else 0)
    colw1 = colw1_arg or (w1[int(len(w1) * 0.9)] if w1 else 0)
    print("column left edges: %s  two-column width ~%.1f (%d lines)  single width ~%.1f (%d lines)" % (col_x, colw2, len(w2), colw1, len(w1)))
    lines.sort(key=lambda r: (r[0], 0 if r[1] == col_x[0] else 1, r[2]))
    grant_norm, grant_mark, refuse, refuse_k = [], [], [], []
    for i, (pno, x0, y0, x1, t, pitch) in enumerate(lines):
        if x0 not in col_x:
            continue
        colw = colw2 if x1 - x0 < 250 else colw1
        if colw <= 0:
            continue
        cells = colw / pitch
        n, pairs, marks = cells_of(t)
        need = n - 0.5 * pairs
        last = t.rstrip()[-1:] if t.rstrip() else ""
        if any(ch.isascii() and (ch.isalnum() or ch == " ") for ch in t.rstrip()):
            continue  # proportional Latin/digits/spaces: cell count unreliable
        if "--single" in sys.argv and (x1 - x0) < 250:
            continue
        if "--double" in sys.argv and (x1 - x0) >= 250:
            continue
        if need > cells + 0.10:
            rec = (pno, y0, t, need - cells, marks, last)
            (grant_mark if (last in MID or last in CLOSE or last in OPEN) else grant_norm).append(rec)
            if show:
                print("GRANT%s p%d y=%.1f over=%.2f marks=%s per=%.2f pitch=%.2f %s" % (
                    "-mark" if last in MID else "-norm", pno, y0, need - cells, "".join(marks),
                    (need - cells) / max(len(marks), 1), pitch, t))
        elif need > cells - 1.02 and i + 1 < len(lines) and lines[i + 1][0] == pno and lines[i + 1][1] == x0 and 0 < lines[i + 1][2] - y0 < 40:
            nxt = lines[i + 1][4].strip()
            if not nxt or t.endswith(" ") or t.endswith("　"):
                continue  # a trailing space marks a paragraph end in the PDF text
            c = nxt[0]
            if c in MID or c in CLOSE:
                continue
            unit = c + nxt[1] if len(nxt) > 1 and (nxt[1] in MID or nxt[1] in CLOSE) else c
            n2, p2, m2 = cells_of(t + unit)
            need2 = n2 - 0.5 * p2
            if need2 > cells + 0.10 and x1 - x0 >= colw - 0.6 * pitch:
                (refuse_k if len(unit) > 1 else refuse).append((pno, y0, t, need2 - cells, m2 if len(unit) > 1 else marks, unit))
                if show and marks:
                    print("REFUSE p%d y=%.1f over=%.2f marks=%s per=%.2f next=%s %s" % (
                        pno, y0, need2 - cells, "".join(marks), (need2 - cells) / len(marks), c, t))

    def table(title, recs):
        print("== %s: %d" % (title, len(recs)))
        hist = Counter()
        for pno, y0, t, over, marks, c in recs:
            per = over / len(marks) if marks else 9.99
            hist[(len(marks), round(per, 1))] += 1
        for (nm, per), cnt in sorted(hist.items()):
            print("   marks=%d per-mark=%.1f : %d" % (nm, per, cnt))
        for pno, y0, t, over, marks, c in recs[:12]:
            print("   p%d over=%.2f marks=%s per=%.2f %s %s" % (pno, over, "".join(marks), over / len(marks) if marks else 0, c, t[:34]))

    def cap_norm(marks):
        return 0.5 if marks else 0.0

    def cap_kinsoku(marks):
        return 1.0 + 0.5 * (len(marks) - 1) if marks else 0.0

    print("== H2: a normal pull-in is at most HALF A CELL per line (from any marks); a kinsoku-final line hangs its mark (1.0) + 0.5 per other mark")
    ok = bad = 0
    for pno, y0, t, over, marks, c in grant_norm:
        if over <= cap_norm(marks) + 0.05:
            ok += 1
        else:
            bad += 1
            print("   GRANT-norm beyond cap: p%d y=%.0f over=%.2f cap=%.2f marks=%s %s" % (pno, y0, over, cap_norm(marks), "".join(marks), t))
    print("   grant-norm: %d agree / %d disagree" % (ok, bad))
    ok = bad = 0
    for pno, y0, t, over, marks, c in grant_mark:
        if over <= cap_kinsoku(marks) + 0.05:
            ok += 1
        else:
            bad += 1
            print("   GRANT-kinsoku beyond cap: p%d y=%.0f over=%.2f cap=%.2f marks=%s %s" % (pno, y0, over, cap_kinsoku(marks), "".join(marks), t))
    print("   grant-kinsoku: %d agree / %d disagree" % (ok, bad))
    ok = bad = 0
    for pno, y0, t, over, marks, c in refuse:
        cap = cap_kinsoku(marks) if (t.rstrip()[-1:] in OPEN) else cap_norm(marks)
        if over > cap - 0.05:
            ok += 1
        else:
            bad += 1
            print("   REFUSE within cap: p%d y=%.0f over=%.2f cap=%.2f marks=%s next=%s %s" % (pno, y0, over, cap, "".join(marks), c, t))
    print("   refuse: %d agree / %d disagree" % (ok, bad))
    # kinsoku-final: the unit X+mark was refused; what would the gap shrink have needed?
    print("== kinsoku-final: granted vs refused, excess = demand - 0.5 x marks (cells), per gap = excess / chars")
    for title, recs in (("GRANTED", grant_mark), ("REFUSED", refuse_k)):
        rows = []
        for pno, y0, t, over, marks, c in recs:
            nchars = sum(1 for ch in t if is_full(ch)) + (1 if title == "REFUSED" else 0)
            excess = over - 0.5 * len(marks)
            tn = t.replace(" ", "").replace("　", "")
            jc = jc_of.get(tn[:8], "?") if jc_of else ""
            rows.append((excess / max(nchars, 1), excess, over, len(marks), "".join(marks), pno, y0, jc, t[:30]))
        rows.sort(reverse=(title == "GRANTED"))
        print("   %s: %d; extreme 10:" % (title, len(rows)))
        for r in rows[:10]:
            print("     per-gap=%+.4f excess=%+.2f over=%.2f marks=%d(%s) p%d y=%.0f jc=%s %s" % r)
        if jc_of:
            byjc = Counter()
            for r in rows:
                byjc[(r[7], "excess>0.05" if r[1] > 0.05 else "excess<=0.05")] += 1
            print("     by jc:", dict(byjc))
    if "--contra" in sys.argv:
        print("== GRANT-norm with per-mark > 0.27")
        for pno, y0, t, over, marks, c in grant_norm:
            if marks and over / len(marks) > 0.27:
                print("   p%d y=%.0f over=%.2f marks=%s per=%.2f %s" % (pno, y0, over, "".join(marks), over / len(marks), t))
        print("== REFUSE with per-mark <= 0.27")
        for pno, y0, t, over, marks, c in refuse:
            if marks and over / len(marks) <= 0.27:
                print("   p%d y=%.0f over=%.2f marks=%s per=%.2f next=%s %s" % (pno, y0, over, "".join(marks), over / len(marks), c, t))
        return
    table("GRANT, line ends in a normal character", grant_norm)
    table("GRANT, line ends in a mid mark", grant_mark)
    table("REFUSE (next char normal, marks on the line)", [r for r in refuse if r[4]])
    print("== REFUSE without marks: %d" % sum(1 for r in refuse if not r[4]))


if __name__ == "__main__":
    main()
