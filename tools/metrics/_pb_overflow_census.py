# -*- coding: utf-8 -*-
"""How often does a rendered line run past the page's right edge?

S1208 fixed WHERE tokyoshugyo's marked item breaks, but not how it is drawn: Word
compresses its two commas by 3.9pt each and ends 0.5pt inside the margin, while Oxi
draws the characters at their natural advances and ends 7.3pt OUTSIDE it. The break
is a cell rule; the drawing is not, and the default build overflows the same way.

Before writing a render-side compressor, count what it would be worth: every text
element whose right edge lands past the section's own right margin, per document.
Floating shapes and text boxes are excluded by only looking at elements that carry a
paragraph index, and a 0.5pt tolerance keeps device rounding out of the count.

    python _pb_overflow_census.py            # every golden document
    python _pb_overflow_census.py tokyo      # id prefix filter
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
EXE = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
TOL = 0.5


def right_edge(docx):
    """Section right edge in pt, from the FIRST sectPr (pgSz w - pgMar right)."""
    try:
        x = zipfile.ZipFile(docx).read("word/document.xml").decode("utf-8")
    except Exception:  # noqa: BLE001
        return None
    m = re.search(r'<w:pgSz [^>]*w:w="(\d+)"', x)
    n = re.search(r'<w:pgMar [^>]*w:right="(\d+)"', x)
    if not m or not n:
        return None
    return (int(m.group(1)) - int(n.group(1))) / 20.0


def dump(docx):
    with tempfile.TemporaryDirectory() as td:
        out = os.path.join(td, "d.json")
        r = subprocess.run([EXE, docx, os.path.join(td, "p"), "96",
                            "--dump-layout=" + out],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if r.returncode != 0 or not os.path.exists(out):
            return None
        return json.load(open(out, encoding="utf-8"))


def main():
    pref = sys.argv[1] if len(sys.argv) > 1 else ""
    files = sorted(f for f in os.listdir(DOCS)
                   if f.endswith(".docx") and not f.startswith("~$")
                   and f.startswith(pref) if not pref or f.startswith(pref))
    rows = []
    for f in files:
        path = os.path.join(DOCS, f)
        edge = right_edge(path)
        if edge is None:
            continue
        d = dump(path)
        if d is None:
            print("%-34s RENDER FAILED" % f[:34])
            continue
        n_over = 0
        worst = 0.0
        worst_txt = ""
        n_text = 0
        for pg in d.get("pages", []):
            for el in pg.get("elements", []):
                if el.get("type") != "text" or el.get("para_idx") is None:
                    continue
                n_text += 1
                over = el["x"] + el.get("w", 0.0) - edge
                if over > TOL:
                    n_over += 1
                    if over > worst:
                        worst, worst_txt = over, el.get("text", "")[:22]
        rows.append((n_over, worst, f[:30], n_text, worst_txt))
        print("%-32s edge %6.2f  over %4d / %6d  worst %6.2fpt  %s"
              % (f[:32], edge, n_over, n_text, worst, worst_txt), flush=True)
    rows.sort(reverse=True)
    print("\ndocuments with any overflow: %d of %d"
          % (sum(1 for r in rows if r[0]), len(rows)))
    print("worst ten by count:")
    for n_over, worst, name, n_text, txt in rows[:10]:
        if not n_over:
            break
        print("  %-32s %4d elements, worst %6.2fpt  %s" % (name, n_over, worst, txt))


if __name__ == "__main__":
    main()
