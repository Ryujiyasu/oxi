# -*- coding: utf-8 -*-
"""Column-aware line comparison: Word PDF vs Oxi --dump-layout.

Row-order comparison of a two-column page interleaves the columns and breaks on
every table, so a differing line looked like a differing paragraph.  This
instrument groups every line by (page, column, y), walks each column top to
bottom, and prints Word and Oxi side by side with the character count of each
line, marking the first line whose text differs.

Usage: python tools/metrics/_col_lines.py <docx> <word.pdf> [--pages=3,4] [--cols=2] [--all]
Environment flags (OXI_S1318=1, OXI_S1336=1 ...) are passed through to the renderer.
"""
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
REPO = Path(__file__).resolve().parents[2]
GDI = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"

docx = sys.argv[1]
pdf = sys.argv[2]
pages = None
ncols = 2
show_all = "--all" in sys.argv
for a in sys.argv[3:]:
    if a.startswith("--pages="):
        pages = [int(x) for x in a.split("=")[1].split(",")]
    if a.startswith("--cols="):
        ncols = int(a.split("=")[1])


def norm(t):
    return t.replace(" ", "").replace("　", "").replace("\t", "")


def word_lines(doc, pno, width):
    """[(col, y, x, text)] for page pno (1-based)."""
    page = doc[pno - 1]
    out = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"])
            if not norm(t):
                continue
            x0, y0, x1, y1 = l["bbox"]
            col = min(int(x0 // (width / ncols)), ncols - 1)
            out.append((col, round(y0, 1), round(x0, 1), t))
    return merge(out)


def merge(items):
    """Join fragments sharing (col, y within 1.5pt) in x order."""
    items.sort(key=lambda r: (r[0], r[1], r[2]))
    lines = []
    for col, y, x, t in items:
        if lines and lines[-1][0] == col and abs(lines[-1][1] - y) < 1.5:
            lines[-1] = (col, lines[-1][1], lines[-1][2], lines[-1][3] + t)
        else:
            lines.append((col, y, x, t))
    return lines


def oxi_dump(docx):
    tmp = tempfile.mkdtemp()
    dump = os.path.join(tmp, "dump.json")
    subprocess.run([str(GDI), docx, os.path.join(tmp, "p"), "--dump-layout=" + dump],
                   capture_output=True, timeout=600)
    return json.load(open(dump, encoding="utf-8"))


def oxi_lines(dump, pno, width):
    out = []
    for page in dump.get("pages", []):
        if page["page"] != pno:
            continue
        # column of a paragraph line = column of its leftmost element (a
        # distributed heading spreads its characters across the page)
        by_line = {}
        for el in page.get("elements", []):
            if el.get("type") != "text" or not norm(el.get("text", "")):
                continue
            key = (el.get("para_idx"), el.get("cell_para_idx"), el.get("cell_row_idx"), el.get("cell_col_idx"), round(el["y"], 1))
            by_line.setdefault(key, []).append(el)
        for key, els in by_line.items():
            els.sort(key=lambda e: e["x"])
            span = els[-1]["x"] - els[0]["x"]
            if span > width / ncols:
                # one line spread over the page (a distributed heading)
                out.append((0, key[4], round(els[0]["x"], 1), "".join(e["text"] for e in els)))
                continue
            # a paragraph continuing from one column into the next has lines
            # at the same y in both: split by column
            groups = {}
            for e in els:
                groups.setdefault(min(int(e["x"] // (width / ncols)), ncols - 1), []).append(e)
            for col, g in groups.items():
                out.append((col, key[4], round(g[0]["x"], 1), "".join(e["text"] for e in g)))
    return merge(out)


def main():
    doc = fitz.open(pdf)
    dump = oxi_dump(docx)
    width = doc[0].rect.width
    plist = pages or list(range(1, len(doc) + 1))
    summary = "--summary" in sys.argv
    for pno in plist:
        w = word_lines(doc, pno, width)
        o = oxi_lines(dump, pno, width)
        if summary:
            cells = []
            for col in range(ncols):
                wl = [norm(r[3]) for r in w if r[0] == col]
                ol = [norm(r[3]) for r in o if r[0] == col]
                first = next((i for i in range(max(len(wl), len(ol))) if (wl[i] if i < len(wl) else "") != (ol[i] if i < len(ol) else "")), None)
                cells.append("col%d W%3d O%3d first-diff=%s" % (col, len(wl), len(ol), "-" if first is None else first))
            print("p%-3d %s" % (pno, " | ".join(cells)))
            continue
        for col in range(ncols):
            wl = [r for r in w if r[0] == col]
            ol = [r for r in o if r[0] == col]
            print("== page %d col %d: word %d lines / oxi %d lines" % (pno, col, len(wl), len(ol)))
            first = None
            for i in range(max(len(wl), len(ol))):
                wt = norm(wl[i][3]) if i < len(wl) else ""
                ot = norm(ol[i][3]) if i < len(ol) else ""
                same = wt == ot
                if not same and first is None:
                    first = i
                if show_all or not same or (first is not None and i <= first + 2):
                    mark = "  " if same else "!!"
                    print("%s %2d W%3d y=%6.1f %-24s | O%3d y=%6.1f %s" % (
                        mark, i, len(wt), wl[i][1] if i < len(wl) else -1, wt[:24],
                        len(ot), ol[i][1] if i < len(ol) else -1, ot[:24]))


if __name__ == "__main__":
    main()
