# -*- coding: utf-8 -*-
"""Sweep the S1167 "clearly under the em" tolerance against Word's own lines.

The mark and the ideograph do not always come from the same width table -- MS
Mincho's measured map has the marks at 7.900 and leaves the ideographs at the
nominal 8.000 -- so the em test needs a tolerance, and the tolerance has to be
read off the documents rather than picked.  Run every scoped document at each
candidate and print the line agreement, so the choice is a table.

    python _cb_tol_sweep.py 1.0 0.95 0.9
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8")

import _cb_budget as B  # noqa: E402
import _cb_font_census as C  # noqa: E402

# The documents the credit actually moves: everything else is byte-identical at
# every tolerance and only costs render time.
MOVERS = ["c7b923e5", "459f05", "a47e6c", "nedocontract", "d77a", "b837",
          "1636d28e", "15076df"]


def main():
    tols = [a for a in sys.argv[1:]] or ["1.0", "0.95", "0.9", "0.85"]
    docs = []
    for f in sorted(os.listdir(B.DOCS)):
        if not f.endswith(".docx") or f.startswith("~$"):
            continue
        if not any(f.startswith(p) for p in MOVERS):
            continue
        if not os.path.exists(os.path.join(B.DOCS, f)[:-5] + "_rt.pdf"):
            continue
        docs.append(os.path.join(B.DOCS, f))
    head = "%-30s %-10s" % ("document", "pre-S1167")
    for t in tols:
        head += "%-10s" % ("tol " + t)
    print(head)
    for path in docs:
        name = os.path.basename(path)[:-5]
        pre = B.match_report(path, "OXI_S1167_DISABLE=1", "pre", quiet=True)
        row = "%-30s %4d/%-5d" % (name[:30], pre[0], pre[1])
        for t in tols:
            r = B.match_report(path, "OXI_S1167_TOL=" + t, "t" + t, quiet=True)
            row += "%4d%-6s" % (r[0], " (%+d)" % (r[0] - pre[0]))
        print(row)


if __name__ == "__main__":
    main()
