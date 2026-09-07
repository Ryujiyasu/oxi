# -*- coding: utf-8 -*-
"""Line-sequence agreement between a Word PDF and Oxi (--dump-layout via
_col_lines.py --all): the two line sequences are aligned with difflib, so one
extra or missing line does not cascade into every line below it.

Usage: python tools/metrics/_pb_agree.py <docx> <word.pdf> [--quiet]
Environment flags (OXI_S1318=1 ...) pass through to the renderer.
"""
import difflib
import os
import re
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
docx, pdf = sys.argv[1], sys.argv[2]
extra = [a for a in sys.argv[3:] if a.startswith("--") and a != "--quiet"]
out = subprocess.run([sys.executable, os.path.join(HERE, "_col_lines.py"), docx, pdf, "--all"] + extra,
                     capture_output=True, text=True, encoding="utf-8", errors="replace").stdout
W, O = [], []
for l in out.splitlines():
    m = re.match(r"^(?:!!)?\s*\d+ W\s+(\d+) y=\s*([-\d.]+) (.*?)\s*\| O\s+(\d+) y=\s*([-\d.]+) (.*)$", l)
    if not m:
        continue
    wn, wy, wt, on, oy, ot = m.groups()
    if int(wn) > 0:
        W.append(wt.strip())
    if int(on) > 0:
        O.append(ot.strip())
sm = difflib.SequenceMatcher(a=W, b=O, autojunk=False)
same = sum(b.size for b in sm.get_matching_blocks())
print("word lines %d, oxi lines %d, identical %d (%.1f%%)" % (len(W), len(O), same, 100.0 * same / max(1, len(W))))
if "--quiet" not in sys.argv:
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        print("-- %s" % tag)
        for t in W[i1:i2]:
            print("   W %2d %s" % (len(t), t))
        for t in O[j1:j2]:
            print("   O %2d %s" % (len(t), t))
