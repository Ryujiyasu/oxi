# -*- coding: utf-8 -*-
"""Put one line's characters side by side, Word's pen against Oxi's.

The aggregates say the per-glyph error is small and the CJK/Latin boundary is a
flat quarter em, but c7b923e5's "API" is 0.76pt wide of Word and three glyphs
cannot hold that. This walks a single line character by character -- Word's
origin from the PDF, Oxi's from `--dump-glyphs` -- so the surplus can be read
off where it is actually spent instead of inferred from a sum.

Both sides are put in the same units: the PDF's 600dpi-snapped size is divided
out and re-applied at the nominal size, so a column of advances can be compared
directly with Oxi's.

    python _cw_latin_line.py c7b923e5 API
    python _cw_latin_line.py d77a58 URL --page 2
"""
import argparse
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import _cb_budget as B                                   # noqa: E402
from _cw_latin_adv import nominal_size, is_latin, is_cjk  # noqa: E402


def word_lines(rt):
    import fitz
    out = []
    for pi, pg in enumerate(fitz.open(rt), 1):
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                chars = []
                for s in ln["spans"]:
                    nom = nominal_size(round(s["size"], 2))
                    k = nom / s["size"]          # undo the 600dpi snap
                    for c in s["chars"]:
                        chars.append((c["c"], c["origin"][0] * k, nom, s["font"]))
                if "".join(c[0] for c in chars).strip():
                    out.append((pi, chars))
    return out


def oxi_lines(docx, tag):
    out = os.path.join(B.OUT, "_cwline_" + tag)
    os.makedirs(B.OUT, exist_ok=True)
    gl = out + "_glyphs.json"
    r = subprocess.run([B.GDI, docx, out, "--dump-glyphs=" + gl],
                       capture_output=True)
    if r.returncode != 0:
        sys.exit(r.stderr.decode("utf-8", "replace")[-1500:])
    data = json.load(open(gl, encoding="utf-8"))
    lines = []
    for pg in data["pages"]:
        rows = {}
        for g in pg["glyphs"]:
            rows.setdefault(round(g["top"], 1), []).append(g)
        for top in sorted(rows):
            gs = sorted(rows[top], key=lambda g: g["x"])
            lines.append((pg["page"], [(g["char"], g["x"], g["font_size"],
                                        g["font_family"]) for g in gs]))
    return lines


def pick(lines, needle, page):
    for pi, chars in lines:
        if page and pi != page:
            continue
        if needle in "".join(c[0] for c in chars):
            return pi, chars
    return None, None


def show(label, chars):
    print("  %s: %s" % (label, "".join(c[0] for c in chars)[:90]))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("prefix")
    ap.add_argument("needle")
    ap.add_argument("--page", type=int, default=0)
    a = ap.parse_args()

    docx = B.docx_for(a.prefix)
    wp, wc = pick(word_lines(docx[:-5] + "_rt.pdf"), a.needle, a.page)
    op, oc = pick(oxi_lines(docx, a.prefix), a.needle, a.page)
    if wc is None or oc is None:
        sys.exit("not found: word=%s oxi=%s" % (wc is not None, oc is not None))
    print("== %s ==  '%s'  word p%d / oxi p%d" % (a.prefix, a.needle, wp, op))
    show("word", wc)
    show("oxi ", oc)

    wt, ot = "".join(c[0] for c in wc), "".join(c[0] for c in oc)
    wi, oi = wt.find(a.needle), ot.find(a.needle)
    lo = min(wi, oi, 6)
    print("\n%-3s %-9s %-9s %-9s %-9s %s"
          % ("ch", "word_adv", "oxi_adv", "o-w", "cum_o-w", "font"))
    cum = 0.0
    n = min(len(wc) - wi, len(oc) - oi) - 1
    for k in range(-lo, n):
        wch, wx, wfs, wfn = wc[wi + k]
        och, ox, _ofs, ofn = oc[oi + k]
        if wch != och:
            print("  ...diverges at %r vs %r" % (wch, och))
            break
        wa = wc[wi + k + 1][1] - wx
        oa = oc[oi + k + 1][1] - ox
        cum += oa - wa
        mark = "" if not (is_cjk(wch) and is_latin(wc[wi + k + 1][0])) else "  <- CJK|lat"
        print("%-3s %-9.3f %-9.3f %-9.3f %-9.3f %s%s"
              % (wch, wa, oa, oa - wa, cum, wfn[:12], mark))


if __name__ == "__main__":
    main()
