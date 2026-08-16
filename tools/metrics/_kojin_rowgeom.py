# -*- coding: utf-8 -*-
"""Row geometry of kojin's page-3 table: Word (PDF rules) vs Oxi (border dump).

The dependency chain recorded at the end of the 2026-08-16 session is

    table row geometry -> kojin's paragraph position (13.89pt) -> the S693
    last-line leniency can be removed

and the last step cannot move until the first one does.  Page 3 shows +-20pt
swings per row and one extra line overall (Word 35 / Oxi 36), so the row heights
themselves are what has to be pinned, not the page-bottom rule.

Word's own row boundaries are the horizontal rules it paints, so read them
straight out of the exported PDF and put them next to Oxi's `border` elements
from --dump-layout.  Both are in points from the page top.

    python _kojin_rowgeom.py word      # export + cache the PDF, list page-3 rules
    python _kojin_rowgeom.py oxi       # Oxi's page-3 border rows
    python _kojin_rowgeom.py cmp       # joined, per-row delta
    python _kojin_rowgeom.py lines     # per-text-line y join for the same page
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_kojin_rowgeom")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
# OXI_DOC=<prefix> points the same join at any corpus document.
DOC = os.environ.get("OXI_DOC", "kojin")
DOCX = next(os.path.join(DOCS, f) for f in sorted(os.listdir(DOCS))
            if f.startswith(DOC) and f.endswith(".docx")
            and not f.startswith("~$"))
PDF = os.path.join(OUT, DOC + ".pdf")
# OXI_ENVS="OXI_S391_PER_LINE_LRPB=0,..." renders a variant into its own cache
ENVS = os.environ.get("OXI_ENVS", "")
TAG = "".join(c for c in ENVS if c.isalnum())[-40:] or "base"
LAYOUT = os.path.join(OUT, "%s_layout_%s.json" % (DOC, TAG))
PAGE = int(os.environ.get("OXI_PAGE", "3"))     # 1-based
TOL = 1.5                                        # pt, rule pairing tolerance


def _ensure_pdf():
    if os.path.exists(PDF):
        return PDF
    os.makedirs(OUT, exist_ok=True)
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
    try:
        d.ExportAsFixedFormat(PDF, 17)
    finally:
        d.Close(False)
        app.Quit()
    return PDF


def _ensure_layout(force=False):
    if os.path.exists(LAYOUT) and not force:
        return LAYOUT
    os.makedirs(OUT, exist_ok=True)
    env = dict(os.environ)
    for kv in [s for s in ENVS.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v if v != "" else "1"
    subprocess.run([GDI, DOCX, os.path.join(OUT, "png_" + TAG),
                    "--dump-layout=" + LAYOUT], check=True, capture_output=True,
                   env=env)
    return LAYOUT


def word_rules(page=PAGE):
    """Horizontal rules Word paints on the page, as y (pt from top)."""
    import fitz
    doc = fitz.open(_ensure_pdf())
    pg = doc[page - 1]
    ys = []
    for d in pg.get_drawings():
        for it in d["items"]:
            if it[0] == "l":
                p0, p1 = it[1], it[2]
                if abs(p0.y - p1.y) < 0.3 and abs(p1.x - p0.x) > 20:
                    ys.append((round((p0.y + p1.y) / 2, 2),
                               round(min(p0.x, p1.x), 1),
                               round(max(p0.x, p1.x), 1)))
            elif it[0] == "re":
                r = it[1]
                if r.height < 2.0 and r.width > 20:
                    ys.append((round((r.y0 + r.y1) / 2, 2),
                               round(r.x0, 1), round(r.x1, 1)))
    return sorted(ys)


def word_lines(page=PAGE):
    """Text lines: (y_top, y_bottom, x0, text)."""
    import fitz
    doc = fitz.open(_ensure_pdf())
    pg = doc[page - 1]
    out = []
    for b in pg.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            txt = "".join(s["text"] for s in ln["spans"])
            if not txt.strip():
                continue
            x0, y0, x1, y1 = ln["bbox"]
            out.append((round(y0, 2), round(y1, 2), round(x0, 2), txt))
    return sorted(out)


def oxi_page(page=PAGE, force=False):
    pages = json.load(open(_ensure_layout(force), encoding="utf-8"))["pages"]
    return pages[page - 1]


def oxi_rules(page=PAGE):
    ys = []
    for e in oxi_page(page)["elements"]:
        if e["type"] != "border":
            continue
        h = e.get("h") or 0.0
        w = e.get("w") or 0.0
        if h <= 2.0 and w > 20:
            ys.append((round(e["y"] + h / 2.0, 2), round(e["x"], 1),
                       round(e["x"] + w, 1)))
    return sorted(ys)


def oxi_lines(page=PAGE):
    rows = {}
    for e in oxi_page(page)["elements"]:
        if e["type"] != "text":
            continue
        key = round(e["y"], 1)
        r = rows.setdefault(key, {"x": e["x"], "t": [], "h": e.get("h") or 0.0,
                                  "off": e.get("text_y_off")})
        r["x"] = min(r["x"], e["x"])
        r["t"].append((e["x"], e.get("text") or ""))
    out = []
    for y, r in rows.items():
        txt = "".join(t for _, t in sorted(r["t"]))
        out.append((y, round(y + r["h"], 2), round(r["x"], 2), txt))
    return sorted(out)


def _pair(a, b, key=lambda v: v[0], tol=TOL):
    """Greedy nearest-neighbour pairing on the first field."""
    out, bi = [], 0
    used = set()
    for x in a:
        best, bj = None, None
        for j, y in enumerate(b):
            if j in used:
                continue
            d = abs(key(y) - key(x))
            if best is None or d < best:
                best, bj = d, j
        if bj is not None and best <= tol:
            used.add(bj)
            out.append((x, b[bj]))
        else:
            out.append((x, None))
    for j, y in enumerate(b):
        if j not in used:
            out.append((None, y))
    return sorted(out, key=lambda p: key(p[0] or p[1]))


def cmp_rules():
    w, o = word_rules(), oxi_rules()
    print("page %d rules: word %d / oxi %d" % (PAGE, len(w), len(o)))
    print("%-9s %-9s %-8s %s" % ("word_y", "oxi_y", "dy", "x span (word|oxi)"))
    for a, b in _pair(w, o, tol=6.0):
        if a and b:
            print("%-9.2f %-9.2f %-8.2f %.0f-%.0f | %.0f-%.0f"
                  % (a[0], b[0], b[0] - a[0], a[1], a[2], b[1], b[2]))
        elif a:
            print("%-9.2f %-9s %-8s %.0f-%.0f" % (a[0], "-", "-", a[1], a[2]))
        else:
            print("%-9s %-9.2f %-8s %s%.0f-%.0f" % ("-", b[0], "-", "", b[1], b[2]))


def _norm(s):
    return "".join(ch for ch in (s or "") if not ch.isspace())


def cmp_lines():
    """Align by TEXT (never by y -- a 1-line slip makes nearest-y lie)."""
    import difflib
    w, o = word_lines(), oxi_lines()
    wt = [_norm(x[3]) for x in w]
    ot = [_norm(x[3]) for x in o]
    print("page %d lines: word %d / oxi %d" % (PAGE, len(w), len(o)))
    print("%-4s %-9s %-9s %-8s %s" % ("#", "word_y", "oxi_y", "dy", "text"))
    i = 0
    sm = difflib.SequenceMatcher(a=wt, b=ot, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                a, b = w[i1 + k], o[j1 + k]
                print("%-4d %-9.2f %-9.2f %-8.2f %s"
                      % (i, a[0], b[0], b[0] - a[0], (a[3] or "")[:40]))
                i += 1
        else:
            for k in range(i1, i2):
                print("%-4d %-9.2f %-9s %-8s W-ONLY %s"
                      % (i, w[k][0], "-", "-", (w[k][3] or "")[:40]))
                i += 1
            for k in range(j1, j2):
                print("%-4d %-9s %-9.2f %-8s O-ONLY %s"
                      % (i, "-", o[k][0], "-", (o[k][3] or "")[:40]))
                i += 1


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "cmp"
    if cmd == "word":
        for y, x0, x1 in word_rules():
            print("%8.2f  %7.1f..%-7.1f" % (y, x0, x1))
    elif cmd == "oxi":
        for y, x0, x1 in oxi_rules():
            print("%8.2f  %7.1f..%-7.1f" % (y, x0, x1))
    elif cmd == "lines":
        cmp_lines()
    elif cmd == "relayout":
        _ensure_layout(force=True)
        print("relaid out")
    else:
        cmp_rules()
