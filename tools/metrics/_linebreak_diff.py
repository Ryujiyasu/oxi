# -*- coding: utf-8 -*-
"""Per-LINE break-divergence finder for ANY corpus document.

The generalized form of `_nedo_linebreak.py` (which is pinned to one JP
contract): export the .docx through Word, read its per-line char stream out of
the PDF, render the same file with Oxi's BODYLINE instrument, align the two
char streams and report every place Oxi's break point differs from Word's.

    OVER-FIT   Oxi keeps a char on the line where Word wrapped
    UNDER-FIT  Oxi wraps where Word kept going (1 Word line -> N Oxi lines)

Word's PDF is cached under pipeline_data/_linebreak/, so re-running after an
Oxi change costs one GDI render.

  python _linebreak_diff.py <docx>                     # first divergences
  python _linebreak_diff.py <docx> all                 # every divergence
  python _linebreak_diff.py <docx> all OXI_S1116_DISABLE=1   # under an env arm
  python _linebreak_diff.py <docx> --refresh           # re-export the PDF

★TRAPS this tool exists to avoid (all three cost a wrong conclusion on
2026-08-14 before being caught):

1. The Oxi dump's width field is `w`, NOT `width`. `e.get("width", 0)` silently
   yields 0 for every text element, so a line's right edge collapses to the
   LAST WORD'S START. That made Oxi look 12.7% narrower than Word and made a
   justified paragraph look unjustified. Both "findings" were the bug.
2. MuPDF's rawdict `lines` are text-run fragments, not rendered lines, and its
   `spans` concatenate chars in CONTENT-STREAM order — this document produced a
   span reading 'rose 9.35down during the coronavirus', spliced out of two
   different rendered lines. Only per-CHAR origins are positional; group those
   by baseline and sort by x.
3. Counting a run's children with a regex over `<w:r>...</w:r>` matches the
   `<w:t>` of runs nested INSIDE the run's own drawing/textbox (the non-greedy
   close tag belongs to the inner run). Use ElementTree direct children.

★★4. A raw Word-vs-Oxi LINE COUNT is not a wrap metric. Word draws the text
   INSIDE embedded charts (axis labels, legends, series names, statistics
   boxes); Oxi draws none of it, and every such label counts as a Word "line".
   On reference__0042471c that is ~54 of the 61-line gap — p7 and p9 lose 22
   lines each to one EMF statistics chart, p2 loses 10 to a DrawingML bar
   chart — leaving only ~7 lines of genuine wrap divergence. Check the
   per-page deficit and look at what is actually missing BEFORE reading a
   line-count gap as a breaking problem.
"""
import difflib
import json
import os
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

import fitz

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
OUT = REPO / "pipeline_data" / "_linebreak"
GDI = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def word_pdf(docx: Path, refresh: bool) -> Path:
    OUT.mkdir(parents=True, exist_ok=True)
    pdf = OUT / (docx.stem + "_word.pdf")
    if pdf.exists() and not refresh:
        return pdf
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(str(docx), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(str(pdf), 17)          # wdExportFormatPDF
    finally:
        d.Close(False)
        app.Quit()
    return pdf


def word_lines(pdf: Path):
    """(page, text) per RENDERED line, in reading order.

    MuPDF's rawdict `lines` are text-run fragments, not rendered lines: Word
    emits justified body text as several show-text chunks, so this document
    came out as 1374 fragments averaging 35.6 chars against Oxi's 756 lines
    averaging 63.8. Group by BASELINE instead — the same grouping the Oxi side
    uses — so both streams are segmented the same way.
    """
    doc = fitz.open(pdf)
    out = []
    for pi in range(doc.page_count):
        d = doc.load_page(pi).get_text("rawdict")
        # PER-CHAR, not per-span: MuPDF concatenates a span's chars in content-
        # stream order, which on this document produced spans whose text is
        # spliced out of two different rendered lines ('rose 9.35down during
        # the coronavirus'). Char origins are positional, so grouping them by
        # baseline and sorting by x reconstructs the true rendered line.
        rows = defaultdict(list)
        for blk in d["blocks"]:
            if blk.get("type") != 0:
                continue
            for ln in blk.get("lines", []):
                for sp in ln.get("spans", []):
                    for c in sp.get("chars", []):
                        if c["c"].strip():
                            rows[round(c["origin"][1], 0)].append((c["origin"][0], c["c"]))
        for y in sorted(rows):
            cs = "".join(ch for _x, ch in sorted(rows[y])).rstrip()
            if cs:
                out.append((pi + 1, cs))
    doc.close()
    return out


def oxi_lines(docx: Path, envs: str):
    """(page, text) per rendered line, from the FULL layout dump.

    The BODYLINE instrument reports body paragraphs only, so comparing it
    against a PDF that also carries table cells and headers mis-aligns the two
    streams wholesale (this document: 587 body lines vs 1374 PDF lines). The
    dump covers everything Oxi draws, which is what the PDF holds.
    """
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    dump = OUT / (docx.stem + "_oxi.json")
    subprocess.run([str(GDI), str(docx), str(OUT / docx.stem), "--dump-layout=" + str(dump)],
                   check=True, capture_output=True, env=env)
    d = json.load(open(dump, encoding="utf-8"))
    out = []
    for pg in d["pages"]:
        rows = defaultdict(list)
        for e in pg["elements"]:
            if (e.get("text") or "").strip():
                rows[round(e.get("y", 0.0), 1)].append(e)
        for y in sorted(rows):
            txt = "".join(e["text"] for e in sorted(rows[y], key=lambda e: e.get("x", 0.0))).rstrip()
            if txt:
                out.append((pg["page"], txt))
    return out


def main():
    args = [a for a in sys.argv[1:]]
    refresh = "--refresh" in args
    args = [a for a in args if a != "--refresh"]
    docx = Path(args[0]).resolve()
    mode = args[1] if len(args) > 1 else "first"
    envs = args[2] if len(args) > 2 else ""

    pdf = word_pdf(docx, refresh)
    wl, ol = word_lines(pdf), oxi_lines(docx, envs)

    # WHITESPACE-INSENSITIVE alignment. The dump carries one element per run,
    # so joining a line's elements drops the inter-word gaps ('1.Introduction',
    # ',whichimpliesthat...'); Word's PDF keeps them. Break points always fall
    # between words, so dropping spaces from BOTH streams costs no divergence
    # signal and stops the aligner from desynchronising on missing spaces.
    wch, wtag, och, otag = [], [], [], []
    for i, (_pg, t) in enumerate(wl):
        for c in t:
            if c.isspace():
                continue
            wch.append(c)
            wtag.append(i)
    for i, (_pg, t) in enumerate(ol):
        for c in t:
            if c.isspace():
                continue
            och.append(c)
            otag.append(i)

    sm = difflib.SequenceMatcher(None, och, wch, autojunk=False)
    pairs = []
    for a, b, size in sm.get_matching_blocks():
        for k in range(size):
            pairs.append((otag[a + k], wtag[b + k]))

    over, prev_w, prev_o = [], None, None
    for o, w in pairs:
        if prev_w is not None and w != prev_w and o == prev_o:
            over.append((prev_w, prev_o))
        prev_w, prev_o = w, o

    w2o = defaultdict(set)
    for o, w in pairs:
        w2o[w].add(o)
    under = [(w, sorted(s)) for w, s in sorted(w2o.items()) if len(s) > 1]

    matched = len(pairs)
    print("%s  |  Word lines %d   Oxi lines %d   matched chars %d/%d"
          % (docx.name, len(wl), len(ol), matched, len(wch)))
    lim = 999 if mode == "all" else 5

    print("\n=== OVER-FIT (Oxi keeps a char past Word's break) ===")
    for w, o in over[:lim]:
        nxt = wl[w + 1][1][:14] if w + 1 < len(wl) else ""
        print("  Wp%-3d Wline%-5d end=%r" % (wl[w][0], w, wl[w][1][-26:]))
        print("        next Word line starts %r" % nxt)
        print("        Oxi line %-5d ...%r" % (o, ol[o][1][-26:] if o < len(ol) else "?"))

    print("\n=== UNDER-FIT (1 Word line -> several Oxi lines) ===")
    for w, os_ in under[:lim]:
        print("  Wp%-3d Wline%-5d -> oxi %s  %r"
              % (wl[w][0], w, os_, wl[w][1][:56]))

    print("\ntotals: overfit=%d  underfit=%d" % (len(over), len(under)))


if __name__ == "__main__":
    main()
