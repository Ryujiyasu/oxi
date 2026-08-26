# -*- coding: utf-8 -*-
"""Which runs does PowerPoint pull back, and by how much of their own width?

A run's drawn width is its /Widths sum PLUS the pen adjustments inside the TJ
array. Most runs carry adjustments that net to nothing (median -0.04% over 295
runs on d15), but a minority are pulled in by 1-3%, and on those the font's
advances alone predict the wrong line break.

What the pullback is NOT (all tested 2026-08-26, d15/d05/d24):
  * not bold      -- d15's deepest are Barlow,Bold 30pt (-2.70%) AND plain
                     non-bold Barlow Light runs (-1.93%, -1.53%)
  * not kerning   -- the face's full GPOS gives -127 units where the page has
                     -214, and -78 where the page has -7
  * not autofit   -- every shape involved is noAutofit
  * not spc       -- the runs carry no letter-spacing, and two runs of ONE
                     paragraph get different treatment
  * not demand    -- runs with most of their box free are pulled as hard as
                     runs that overflow
It is also not uniform within a run: the adjustments scatter (-11..+29
thousandths) rather than repeating one value per character.

Deck-dependent: d05's worst run is -0.38% while d15 has many past -1.5%.

Usage: python tools/metrics/pptx_tj_census.py d15 d05 d24
"""
import glob
import re
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BACKSLASH = chr(92)


def parse_tj(body):
    txt, adj, i = [], [], 0
    while i < len(body):
        ch = body[i]
        if ch == "(":
            depth, i, buf = 1, i + 1, []
            while i < len(body) and depth:
                c = body[i]
                if c == BACKSLASH:
                    buf.append(body[i + 1])
                    i += 2
                    continue
                if c == "(":
                    depth += 1
                elif c == ")":
                    depth -= 1
                    if not depth:
                        i += 1
                        break
                buf.append(c)
                i += 1
            txt.append("".join(buf))
        elif ch == "<":
            j = body.find(">", i)
            if j < 0:
                break
            h = re.sub(r"[^0-9A-Fa-f]", "", body[i + 1 : j])
            if len(h) % 2:
                h += "0"
            try:
                txt.append(bytes.fromhex(h).decode("latin-1"))
            except ValueError:
                pass
            i = j + 1
        elif ch == "-" or ch.isdigit() or ch == ".":
            m = re.match(r"-?\d+\.?\d*", body[i:])
            adj.append(float(m.group()))
            i += len(m.group())
        else:
            i += 1
    return "".join(txt), adj


def widths_for(doc, page):
    """resource name (F1..) -> (FirstChar, [widths])"""
    out = {}
    for xref, _, _, base, name, _, _ in page.get_fonts(full=True):
        obj = doc.xref_object(xref)
        fc = re.search(r"/FirstChar\s+(\d+)", obj)
        wm = re.search(r"/Widths\s+(\d+)\s+0\s+R", obj)
        body = None
        if wm:
            body = doc.xref_object(int(wm.group(1)))
        else:
            inline = re.search(r"/Widths\s*\[(.*?)\]", obj, re.S)
            body = inline.group(1) if inline else None
        if fc and body:
            out[name] = (int(fc.group(1)),
                         [float(x) for x in re.findall(r"-?\d+\.?\d*", body)], base)
    return out


def main():
    decks = sys.argv[1:] or ["d15"]
    for deck in decks:
        pdfs = sorted(glob.glob(f"pipeline_data/pptx_benchmark/dev/pdf/{deck}__*.pdf"))
        if not pdfs:
            print(f"{deck}: no reference pdf")
            continue
        doc = pymupdf.open(pdfs[0])
        rows = []
        for pno in range(len(doc)):
            page = doc[pno]
            wtab = widths_for(doc, page)
            raw = page.read_contents().decode("latin-1", "replace")
            cur = None
            for m in re.finditer(r"/(\w+)\s+([\d.]+)\s+Tf|\[(.*?)\]\s*TJ", raw, re.S):
                if m.group(1):
                    cur = (m.group(1), float(m.group(2)))
                    continue
                if cur is None or cur[0] not in wtab:
                    continue
                txt, adj = parse_tj(m.group(3))
                t = txt.strip()
                if len(t) < 10:
                    continue
                first, widths, base = wtab[cur[0]]
                tot = 0.0
                ok = True
                for c in txt:
                    code = ord(c)
                    if not (first <= code < first + len(widths)):
                        ok = False
                        break
                    tot += widths[code - first]
                if not ok or tot <= 0:
                    continue
                design = tot / 1000 * cur[1]
                pull = -sum(adj) / 1000 * cur[1]
                rows.append((pull / design, pull, design, pno + 1, base, cur[1], txt))
        rows.sort()
        print(f"\n=== {deck}: {len(rows)} runs")
        print(f"{'pull%':>7}{'pull_pt':>9}{'width':>9}{'p#':>5}  font / size / text")
        for r in rows[:8]:
            print(f"{r[0]*100:7.2f}{r[1]:9.2f}{r[2]:9.2f}{r[3]:5}  {r[4].split('+')[-1]} {r[5]:g}  {r[6][:40]!r}")
        mid = rows[len(rows) // 2]
        print(f"  median pull {mid[0]*100:+.2f}%   {mid[6][:40]!r}")


if __name__ == "__main__":
    main()
