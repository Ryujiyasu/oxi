# -*- coding: utf-8 -*-
"""Read the `_pb_vertcols` arms: Word COM truth, Oxi, or both side by side.

Reports, per arm, the (page, band, column) each SECTION occupies. The column
index is read off x -- the arms pin every column to an exact 18.0pt pitch, so
col_k = round((742.65 - x)/18). Word is walked per CHARACTER: a paragraph that
spans several columns reports only its first from `Paragraphs`, and reading
paragraph starts alone shows a staircase that is not there.

    python tools/metrics/_pb_vertcols_read.py word
    python tools/metrics/_pb_vertcols_read.py oxi
    python tools/metrics/_pb_vertcols_read.py both
"""
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
GDI = os.environ.get("OXI_GDI_EXE") or str(
    REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe")
SRC = Path(r"C:\tmp\pb_vertcols")
RIGHT = 742.65      # page 841.9 - right margin 99.25
PITCH = 18.0

wdActiveEndSectionNumber = 2
wdActiveEndPageNumber = 3
wdHorizontalPositionRelativeToPage = 5
wdVerticalPositionRelativeToPage = 6


def col_of(x):
    return round((RIGHT - x) / PITCH)


def word_arm(app, path):
    doc = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
    try:
        # A column is where x CHANGES, not where an x is first seen: bands reuse
        # the same x lower down the page, so a (page, x) key silently collapses
        # bands 1 and 2 onto band 0 and the balance becomes unreadable.
        rows = []
        prev = None
        for p in doc.Paragraphs:
            rng = p.Range
            for pos in range(rng.Start, rng.End):
                c = doc.Range(pos, pos + 1)
                t = c.Text or ""
                if t in ("\r", "\x07", ""):
                    continue
                cr = doc.Range(pos, pos)
                pg = int(cr.Information(wdActiveEndPageNumber))
                x = round(float(cr.Information(wdHorizontalPositionRelativeToPage)), 2)
                if prev == (pg, x):
                    continue
                prev = (pg, x)
                rows.append({
                    "page": pg, "col": col_of(x), "x": x,
                    "sec": int(cr.Information(wdActiveEndSectionNumber)),
                    "y": round(float(cr.Information(wdVerticalPositionRelativeToPage)), 2),
                    "ch": t,
                })
        return rows
    finally:
        doc.Close(False)


def oxi_arm(path):
    with tempfile.TemporaryDirectory(prefix="vc_") as t:
        dj = os.path.join(t, "l.json")
        r = subprocess.run([GDI, str(path), os.path.join(t, "p"), "--dump-layout=" + dj],
                           capture_output=True, timeout=180)
        if r.returncode != 0 or not os.path.exists(dj):
            return None
        dump = json.load(open(dj, encoding="utf-8"))
    out = []
    for pi, pg in enumerate(dump["pages"], 1):
        for e in pg["elements"]:
            if e.get("type") != "text" or not e.get("text"):
                continue
            out.append({"page": pi, "col": col_of(e["x"]), "x": round(e["x"], 2),
                        "sec": None, "y": round(e["y"], 2), "ch": e["text"][0]})
    return out


def show(label, rows):
    if rows is None:
        print("    %s: RENDER FAIL" % label)
        return
    bands = sorted({r["y"] for r in rows})
    print("    %-5s bands_y=%s" % (label, [round(b, 2) for b in bands]))
    for pg in sorted({r["page"] for r in rows}):
        for y in bands:
            cells = sorted([r for r in rows if r["page"] == pg and abs(r["y"] - y) < 0.6],
                           key=lambda r: r["col"])
            if not cells:
                continue
            desc = " ".join("c%d%s%s" % (c["col"], "/s%d" % c["sec"] if c["sec"] else "",
                                         ":" + c["ch"]) for c in cells)
            print("      p%d y=%-7.2f %s" % (pg, y, desc))


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "both"
    arms = sorted(SRC.glob("*.docx"))
    if not arms:
        print("no arms; run _pb_vertcols_gen.py first")
        return
    app = None
    if mode in ("word", "both"):
        import win32com.client as win32
        app = win32.gencache.EnsureDispatch("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
    try:
        for a in arms:
            print("=== %s" % a.stem)
            if mode in ("word", "both"):
                show("word", word_arm(app, a))
            if mode in ("oxi", "both"):
                show("oxi", oxi_arm(a))
    finally:
        if app:
            app.Quit()


if __name__ == "__main__":
    main()
