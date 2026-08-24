# -*- coding: utf-8 -*-
"""S1210: what advance does Word give a full-width glyph under a docGrid?

Two questions at once -- does a font SMALLER than the grid default expand to the
grid pitch (S141 says no), and is the pitch ADDITIVE (fs + charSpace/4096) or
PROPORTIONAL (fs * pitch / default_fs, what layout/mod.rs used)?

The gate in `layout/mod.rs` (S466/S141) says no -- `font_size < default_fs`
disables grid expansion. This measures the question directly on Word's own PDF:
for every full-width CJK glyph, the advance to the NEXT glyph on the same line.

Two traps. Word's PDF export is a 600dpi DEVICE rendering: sizes and glyph
origins are whole 0.12pt pixels, so a single advance is quantised (a 9.3547pt
pitch shows up as 9.36 95% of the time and 9.24 the rest) and the reported span
size is not the run's size (10.5pt -> 10.56). Average over a long run and snap
the size back onto the sizes the document actually uses.
And: a JUSTIFIED line stretches its advances, so a wide advance proves
nothing. So each line is classified first -- a line whose right edge falls short
of the widest line sharing its x-origin (its column) by more than one em cannot
have been stretched, and its advances are the pure pitch.

    python _s1210_pitch.py [doc_id ...]      # default: the charSpace=1453 family
"""
import collections
import os
import re
import sys
import zipfile

import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
OUT = os.path.join(REPO, "pipeline_data", "_s1210")
DEFAULT = ["a1d6e4efa2e7_tokumei_08_01-4", "6514f214e482_tokumei_08_01-2",
           "de6e32e4ba0b_tokumei_08_01-3", "d4d126dfe1d9_tokumei_08_01-1"]


def docx_path(did):
    for f in os.listdir(DOCS):
        if f.startswith(did) and f.endswith(".docx"):
            return os.path.join(DOCS, f)
    return None


def grid_of(path):
    x = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    m = re.search(r'<w:docGrid[^>]*w:charSpace="(-?\d+)"', x)
    return int(m.group(1)) if m else 0


def export(path, pdf):
    """Word COM export. Only when the PDF is missing -- Word is slow."""
    if os.path.exists(pdf):
        return pdf
    import win32com.client
    os.makedirs(os.path.dirname(pdf), exist_ok=True)
    w = win32com.client.Dispatch("Word.Application")
    w.Visible = False
    d = w.Documents.Open(os.path.abspath(path), ReadOnly=True)
    d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
    d.Close(0)
    w.Quit()
    return pdf


def lines_of(pdf):
    doc = fitz.open(pdf)
    for pno in range(doc.page_count):
        for b in doc[pno].get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                ch = []
                for s in l["spans"]:
                    for c in s["chars"]:
                        ch.append((c["origin"][0], c["bbox"][2], round(s["size"], 2),
                                   c["bbox"][2] - c["bbox"][0]))
                ch.sort()
                if ch:
                    yield pno, ch


PX = 0.12  # one 600dpi device pixel, the unit Word's PDF export quantises to


def sizes_used(path):
    x = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    y = zipfile.ZipFile(path).read("word/styles.xml").decode("utf-8")
    out = {int(v) / 2.0 for v in re.findall(r'<w:sz w:val="(\d+)"', x + y)}
    return sorted(out | {10.5})


def snap(reported, used):
    """The PDF's span size is the run's size rounded to a device pixel."""
    best = min(used, key=lambda fs: abs(round(fs / PX) * PX - reported))
    return best if abs(round(best / PX) * PX - reported) < PX / 2 else None


def main(ids):
    for did in ids:
        path = docx_path(did)
        if path is None:
            print("%-34s NO DOCX" % did)
            continue
        cs = grid_of(path) / 4096.0
        used = sizes_used(path)
        pdf = export(path, os.path.join(OUT, did + ".pdf"))
        rows = list(lines_of(pdf))
        widest = collections.defaultdict(float)
        for _, ch in rows:
            widest[round(ch[0][0])] = max(widest[round(ch[0][0])], ch[-1][1])
        runs = collections.defaultdict(list)
        for _, ch in rows:
            if widest[round(ch[0][0])] - ch[-1][1] < ch[-1][2]:
                continue                       # may have been justified -- skip
            run = []
            for c in list(ch) + [None]:
                full = c is not None and abs(c[3] - c[2]) <= 0.5
                if full and (not run or c[2] == run[0][2]):
                    run.append(c)
                    continue
                if len(run) >= 6:
                    runs[run[0][2]].append((run[-1][0] - run[0][0], len(run) - 1))
                run = [c] if full else []
        default_fs = 10.5
        print(chr(10) + "%s   charSpace = %+.4f pt/char" % (did, cs))
        print("   run fs  n_runs  n_adv   measured   additive  proportional   verdict")
        for rep in sorted(runs):
            fs = snap(rep, used)
            if fs is None:
                continue
            span = sum(d for d, _ in runs[rep])
            n = sum(k for _, k in runs[rep])
            if n < 20:
                continue
            meas = span / n
            add = fs + cs
            prop = fs * (default_fs + cs) / default_fs
            verdict = ("ADDITIVE" if abs(meas - add) < abs(meas - prop) - 0.005
                       else "proportional" if abs(meas - prop) < abs(meas - add) - 0.005
                       else "tie")
            print("   %5.2f  %6d  %5d   %8.4f  %9.4f  %12.4f   %s%s"
                  % (fs, len(runs[rep]), n, meas, add, prop, verdict,
                     "  <-- fs < default" if fs < default_fs else ""))


if __name__ == "__main__":
    main(sys.argv[1:] or DEFAULT)
