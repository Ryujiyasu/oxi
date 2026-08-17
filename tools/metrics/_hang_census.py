# -*- coding: utf-8 -*-
"""Which characters does each engine let hang past the right text edge?

The corpus floor converged on one signature: where Word pushes a character to the
next line, Oxi pulls it in and the line ends past the content right edge
(c7b923e5 p3 by 7.45pt, tokyoshugyo p21 by 2.55pt). The hypothesis to test is
"only 約物 may hang". Count it instead of assuming: for every line of every page,
compare the line's right edge with the section's text right edge and tabulate the
LAST character of the lines that exceed it -- separately for Word and Oxi.

    python _hang_census.py c7b923e5 tokyoshugyo albalunaTaidan_6pt
"""
import collections
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8")

# 約物 = the characters JIS X 4051 lets hang / compress. Kept deliberately wide:
# the point is to see whether the exceeding characters fall inside it at all.
YAKUMONO = set("。、．，」』）〕】》〉”’!?！？：；・…―‐-")


def text_right(docx_path):
    z = zipfile.ZipFile(docx_path)
    x = z.read("word/document.xml").decode("utf-8")
    m = re.search(r'<w:pgSz[^>]*w:w="(\d+)"', x)
    r = re.search(r'<w:pgMar[^>]*w:right="(\d+)"', x)
    if not m or not r:
        return None
    return (int(m.group(1)) - int(r.group(1))) / 20.0


def main():
    import _kojin_rowgeom as K
    for doc in sys.argv[1:]:
        os.environ["OXI_DOC"] = doc
        for mod in ("_kojin_rowgeom",):
            if mod in sys.modules:
                del sys.modules[mod]
        import _kojin_rowgeom as K  # noqa: F811  (re-import with the new OXI_DOC)
        # ★Do NOT use pgSz-pgMar as the edge: a multi-section document has a
        # per-section margin (and tokyoshugyo has landscape/table sections), so
        # one global number mixes pages that legitimately run wider. Use WORD's
        # OWN widest line on each page as that page's text edge, and ask how far
        # Oxi exceeds it. The measure is then Word-relative and geometry-free.
        edge = text_right(K.DOCX)
        print("== %s == (declared right edge %.2f; per-page edge = Word's widest line)"
              % (doc, edge or -1))
        import fitz
        pdf = fitz.open(K._ensure_pdf())
        wc, oc = collections.Counter(), collections.Counter()
        wover, oover = [], []
        page_edge = {}
        for pi in range(pdf.page_count):
            rights = []
            for b in pdf[pi].get_text("dict")["blocks"]:
                for ln in b.get("lines", []):
                    t = "".join(s["text"] for s in ln["spans"]).rstrip()
                    if t:
                        rights.append((ln["bbox"][2], t[-1]))
            if not rights:
                continue
            e = max(r for r, _ in rights)
            page_edge[pi + 1] = e
            for r, c in rights:
                if r > e - 0.5:
                    wc[c] += 1
                    wover.append((pi + 1, round(r - e, 2), c))
        for pi in range(pdf.page_count):
            try:
                pg = K.oxi_page(pi + 1)
            except IndexError:
                break
            # ★Group by (cell, y), never y alone: a table row's cells all share
            # a y, so a y-only grouping welds them into one very wide "line" and
            # invents overhangs. (Fifth time this trap has fired today.)
            rows = {}
            for e in pg["elements"]:
                if e["type"] == "text":
                    key = (e.get("cell_row_idx"), e.get("cell_col_idx"),
                           round(e["y"], 1))
                    rows.setdefault(key, []).append(e)
            for _key, v in rows.items():
                es = sorted(v, key=lambda e: e["x"])
                t = "".join(e.get("text") or "" for e in es).rstrip()
                if not t:
                    continue
                x1 = es[-1]["x"] + (es[-1].get("w") or 0)
                e = page_edge.get(pi + 1)
                if e is not None and x1 > e + 0.5:
                    oc[t[-1]] += 1
                    oover.append((pi + 1, round(x1 - e, 2), t[-1]))
        print("  WORD lines AT its own widest: %d  %s"
              % (sum(wc.values()), wc.most_common(8)))
        print("     yakumono %d / other %d"
              % (sum(n for c, n in wc.items() if c in YAKUMONO),
                 sum(n for c, n in wc.items() if c not in YAKUMONO)))
        print("  OXI  lines PAST Word's widest: %d  %s"
              % (sum(oc.values()), oc.most_common(8)))
        print("     yakumono %d / other %d"
              % (sum(n for c, n in oc.items() if c in YAKUMONO),
                 sum(n for c, n in oc.items() if c not in YAKUMONO)))
        if oover:
            print("     OXI worst overhangs:",
                  sorted(oover, key=lambda r: -r[1])[:5])
        if wover:
            print("     WORD worst overhangs:",
                  sorted(wover, key=lambda r: -r[1])[:5])


if __name__ == "__main__":
    main()
