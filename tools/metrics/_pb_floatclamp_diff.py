# -*- coding: utf-8 -*-
"""Which elements did S1268 move, in one doc?  (A = clamp on, B = clamp off)

Renders both arms with the SAME binary, serially, and prints every element
whose position differs.  Use it to tell a float the clamp was CORRECTING from
a float the clamp was HIDING: the clamp can only ever push a box up or left,
so an element that lands closer to Word after the clamp is an element whose
anchor Oxi resolved wrong in the first place.

Usage: _pb_floatclamp_diff.py <docx> [<docx> ...]
"""
import json
import os
import subprocess
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DW = os.path.join(REPO, "tools", "oxi-dwrite-renderer", "target", "release",
                  "oxi-dwrite-renderer.exe")


def dump(docx, disable, out):
    env = dict(os.environ)
    if disable:
        env["OXI_S1268_DISABLE"] = "1"
    else:
        env.pop("OXI_S1268_DISABLE", None)
    subprocess.run([DW, docx, out[:-5], "110", "--dump-layout=" + out],
                   capture_output=True, env=env)
    with open(out, encoding="utf-8") as f:
        return json.load(f)


for docx in sys.argv[1:]:
    tmp = tempfile.mkdtemp(prefix="fcdiff_")
    a = dump(docx, True, os.path.join(tmp, "a.json"))
    b = dump(docx, False, os.path.join(tmp, "b.json"))
    name = os.path.basename(docx)
    if len(a["pages"]) != len(b["pages"]):
        print("%-28s PAGE COUNT %d -> %d" % (name, len(a["pages"]), len(b["pages"])))
    moved = 0
    for pi, (pa, pb) in enumerate(zip(a["pages"], b["pages"])):
        if len(pa["elements"]) != len(pb["elements"]):
            print("%-28s p%-2d element count %d -> %d"
                  % (name, pi + 1, len(pa["elements"]), len(pb["elements"])))
            continue
        for ea, eb in zip(pa["elements"], pb["elements"]):
            if abs(ea["x"] - eb["x"]) > 0.01 or abs(ea["y"] - eb["y"]) > 0.01:
                moved += 1
                if moved <= 12:
                    print("%-28s p%-2d %-6s (%8.2f,%8.2f) -> (%8.2f,%8.2f)  d=(%+7.2f,%+7.2f) %r"
                          % (name, pi + 1, ea["type"], ea["x"], ea["y"], eb["x"], eb["y"],
                             eb["x"] - ea["x"], eb["y"] - ea["y"],
                             (ea.get("text") or "")[:16]))
    print("%-28s page=%.1fx%.1f  moved=%d" % (name, a["pages"][0]["width"],
                                              a["pages"][0]["height"], moved))
