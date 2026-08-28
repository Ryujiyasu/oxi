# -*- coding: utf-8 -*-
"""BARE-control y diagnosis (R33 follow-up).

R33 found the keepNext probe's BARE *control* itself disagrees with Word by one
filler line: Word moves R3 to p2 at fill=55, Oxi at fill=54; Word SPLITs the tall
row at 56-58, Oxi only at 58.  Before touching the keepNext chain, pin what the
one-line gap is made of: page bottom limit, row heights, or the fit test.

Prints Word (PDF truth) vs Oxi (--dump-layout) y for every line of one arm.
Usage: _pb_kntbl_ydiag.py [tag ...]     (default BARE46)
"""
import os, sys, json, subprocess
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pb_kntbl_read import word_pages, oxi_pages
from _pb_kntbl_gen import OUT

tags = sys.argv[1:] or ["BARE46"]
for tag in tags:
    docx = os.path.join(OUT, tag + ".docx")
    w = word_pages(docx)
    o = oxi_pages(docx)
    print(f"\n===== {tag} =====")
    print(f"{'page':>4} {'Word y':>9} {'Oxi y':>9} {'d':>7}  text")
    # match by text, in order
    wi = {t: (p, y) for p, y, t in w}
    seen = set()
    for p, y, t in o:
        key = t.strip()
        wp, wy = wi.get(key, (None, None))
        d = f"{y-wy:+.2f}" if wy is not None and wp == p else ("PAGE" if wy is not None else "  --")
        if key in seen: continue
        seen.add(key)
        print(f"{p:>4} {wy if wy is not None else float('nan'):>9} {y:>9.2f} {d:>7}  {key[:46]}")
    # tail: word lines Oxi never emitted
    for p, y, t in w:
        if t.strip() not in seen:
            print(f"{p:>4} {y:>9} {'--':>9} {'MISS':>7}  {t[:46]}")
