# -*- coding: utf-8 -*-
"""Read the TJ arrays PowerPoint emits, with the pen adjustments inside them.

This is how a line can be drawn NARROWER than the sum of its font's /Widths:
the adjustments pull the pen back. d15 p2 is the case -- the run marked b="1"
carries +214 thousandths (-2.57pt) while an ordinary line of the same length
carries -0.08pt, and that 2.57pt is what lets PowerPoint keep "will" on a line
whose /Widths sum (275.40pt) exceeds its 273.47pt box.

PDF strings come in two spellings -- literal `(...)` and hex `<...>`. Handling
only the hex form finds nothing here and invites the conclusion that the page
carries no positioning adjustments at all, which is how the squeeze finding got
retracted in error.
"""
import glob
import re
import sys

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BACKSLASH = chr(92)


def parse_tj(body):
    """-> (text, adjustments) for one TJ array body."""
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
            j = body.index(">", i)
            txt.append(bytes.fromhex(re.sub(r"\s", "", body[i + 1 : j])).decode("latin-1"))
            i = j + 1
        elif ch == "-" or ch.isdigit() or ch == ".":
            m = re.match(r"-?\d+\.?\d*", body[i:])
            adj.append(float(m.group()))
            i += len(m.group())
        else:
            i += 1
    return "".join(txt), adj


def main():
    p = sorted(glob.glob("pipeline_data/pptx_benchmark/dev/pdf/d15__*.pdf"))[0]
    doc = pymupdf.open(p)
    pg = doc[1]
    raw = pg.read_contents().decode("latin-1", "replace")
    cur = None
    found = 0
    for m in re.finditer(r"/(\w+)\s+([\d.]+)\s+Tf|\[(.*?)\]\s*TJ", raw, re.S):
        if m.group(1):
            cur = (m.group(1), float(m.group(2)))
            continue
        txt, adj = parse_tj(m.group(3))
        if "ownload" not in txt and "Click on the button" not in txt:
            continue
        found += 1
        # A positive TJ number moves the pen BACK by that many thousandths of em.
        effect = -sum(adj) / 1000 * cur[1]
        print(f"font=/{cur[0]} size={cur[1]}")
        print(f"  text = {txt[:70]!r}")
        print(f"  adjustments: n={len(adj)} sum={sum(adj):+.1f} thousandths -> {effect:+.3f}pt")
        print(f"  distinct values: {sorted(set(adj))[:14]}")
    if not found:
        print("no TJ array contained the line -- text may be split across ops")
        # fall back: report every TJ with its font and running text
        cur = None
        for m in re.finditer(r"/(\w+)\s+([\d.]+)\s+Tf|\[(.*?)\]\s*TJ", raw, re.S):
            if m.group(1):
                cur = (m.group(1), float(m.group(2)))
                continue
            txt, adj = parse_tj(m.group(3))
            if txt.strip():
                print(f"  /{cur[0]} {cur[1]:>5}  adj_sum={sum(adj):+8.1f}  {txt[:56]!r}")


if __name__ == "__main__":
    main()
