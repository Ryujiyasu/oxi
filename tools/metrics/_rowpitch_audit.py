# -*- coding: utf-8 -*-
"""Per-ROW pitch audit: group a document's horizontal rules into tables and
compare Word's row pitches against Oxi's, row by row.

`_tblfoot_audit.py` anchors on text, which drifts once a row wraps differently;
this one anchors on the RULES themselves, so each row's box is measured directly
and a per-row error cannot hide inside a multi-row span.

  python tools/metrics/_rowpitch_audit.py <docx> [FLAG=1] [--limit=N]
"""
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
DOCX = Path(sys.argv[1]).resolve()
FLAG = next((a for a in sys.argv[2:] if not a.startswith("--")), None)
LIMIT = next((int(a.split("=")[1]) for a in sys.argv if a.startswith("--limit=")), 12)
PDF = Path(tempfile.gettempdir()) / (DOCX.name + ".truth.pdf")


def merge(ys):
    out = []
    for y in sorted(ys):
        if not out or y - out[-1] > 1.6:
            out.append(y)
    return out


def word_rules():
    if not PDF.exists() or "--reexport" in sys.argv:
        import win32com.client as win32
        w = win32.DispatchEx("Word.Application")
        w.Visible = False
        try:
            d = w.Documents.Open(str(DOCX), ReadOnly=True)
            d.ExportAsFixedFormat(str(PDF), 17)
            d.Close(False)
        finally:
            w.Quit()
    import fitz
    doc = fitz.open(PDF)
    out = []
    for pi in range(doc.page_count):
        ys = {round(d["rect"].y0, 2) for d in doc[pi].get_drawings()
              if d["rect"].height < 8 and d["rect"].width > 40}
        out += [(pi, y) for y in merge(ys)]
    return out


def oxi_rules(flag):
    exe = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
    tmp = Path(tempfile.mkdtemp())
    dump = tmp / "d.json"
    env = dict(os.environ)
    for k in ("OXI_S1191", "OXI_S1189", "OXI_S1188"):
        env.pop(k, None)
    if flag:
        k, _, v = flag.partition("=")
        env[k] = v or "1"
    subprocess.run([str(exe), str(DOCX), str(tmp / "p"), "110", "--dump-layout=%s" % dump],
                   check=True, capture_output=True, env=env)
    import json
    d = json.load(open(dump, encoding="utf-8"))
    out = []
    for pi, pg in enumerate(d["pages"]):
        ys = {round(e["y"], 2) for e in pg["elements"]
              if e["type"] == "border" and (e.get("w") or 0) > 40}
        out += [(pi, y) for y in merge(ys)]
    return out


def groups(rules, gap=40.0):
    """Split a page's rule list into tables: a gap > `gap` starts a new one."""
    out = []
    cur = []
    for pi, y in rules:
        if cur and (cur[-1][0] != pi or y - cur[-1][1] > gap):
            out.append(cur)
            cur = []
        cur.append((pi, y))
    if cur:
        out.append(cur)
    return [g for g in out if len(g) >= 2]


def pitches(g):
    return [round(g[k + 1][1] - g[k][1], 2) for k in range(len(g) - 1)]


W = groups(word_rules())
A = groups(oxi_rules(None))
B = groups(oxi_rules(FLAG)) if FLAG else None
print("word tables %d / oxi tables %d" % (len(W), len(A)))
print("%-4s %-5s %-7s %s" % ("grp", "page", "rows", "per-row pitch  (word | oxi | d)"))


def nearest(g, pool):
    """★Match groups by POSITION, never by index — the two sides split into a
    different number of tables (25 vs 23 here) and an index-wise pairing
    compares unrelated tables (trap #41)."""
    pi, y = g[0]
    cand = [h for h in pool if h[0][0] == pi]
    if not cand:
        return None
    return min(cand, key=lambda h: abs(h[0][1] - y))


shown = 0
for i, gw in enumerate(W):
    if shown >= LIMIT:
        break
    ga = nearest(gw, A)
    if ga is None or abs(ga[0][1] - gw[0][1]) > 25.0:
        continue
    shown += 1
    pw, pa = pitches(gw), pitches(ga)
    i = i
    if len(pw) != len(pa):
        print("%-4d %-5d ROW COUNT DIFFERS  word %d / oxi %d  (top w %.1f / o %.1f)"
              % (i, gw[0][0], len(pw), len(pa), gw[0][1], ga[0][1]))
        print("      word %s" % pw)
        print("      oxi  %s" % pa)
        continue
    ds = [round(a - w, 2) for w, a in zip(pw, pa)]
    tot = round(sum(ds), 2)
    flagged = " <<<" if abs(tot) > 0.5 else ""
    print("%-4d %-5d %-7d sum_d %+6.2f  d=%s%s" % (i, gw[0][0], len(pw), tot, ds, flagged))
    gb = nearest(gw, B) if B else None
    if gb:
        pb = pitches(gb)
        if len(pb) == len(pw):
            print("       (flag) sum_d %+6.2f  d=%s"
                  % (round(sum(b - w for w, b in zip(pw, pb)), 2),
                     [round(b - w, 2) for w, b in zip(pw, pb)]))
