# -*- coding: utf-8 -*-
"""Why does a PROPORTIONAL Japanese face lay out at a full em in c7b923e5?

Measured on that document's unjustified lines, MS PGothic's の advances
0.987-1.001em, while the face's own design table says 0.816em. Two readings
survive: Word is substituting a fullwidth face for those runs (against it: on
the same page MuPDF labels the heading spans MS-Gothic and the body spans
MS-PGothic, so the labels do distinguish faces), or the section's docGrid puts
CJK on a full-em pitch regardless of the face.

The grid is the testable half. Same sentence, same face, left-aligned so no line
is ever stretched, under each docGrid setting:

    python _cb_grid_probe.py          # generate, export through Word, measure
    python _cb_grid_probe.py --keep   # reuse the PDFs already exported
"""
import collections
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8")

import _cb_budget as B  # noqa: E402
import _cb_gen as G  # noqa: E402
import _cb_pgothic_adv as A  # noqa: E402

OUT = os.path.join(B.REPO, "pipeline_data", "_cb_grid")
ARMS = [
    ("none", "ＭＳ Ｐゴシック"),
    ("lines", "ＭＳ Ｐゴシック"),
    ("linesAndChars", "ＭＳ Ｐゴシック"),
    ("lines", "ＭＳ ゴシック"),          # control: a fullwidth face, same grid
]


def export(docx):
    pdf = docx[:-5] + "_rt.pdf"
    if os.path.exists(pdf):
        return pdf
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx, ReadOnly=True)
    try:
        d.ExportAsFixedFormat(pdf, 17)
    finally:
        d.Close(False)
        app.Quit()
    return pdf


def measure(pdf):
    """Median advance per character, over CJK-CJK pairs inside one span.

    Every line is usable here: the probe is left-aligned, so Word never
    stretches one."""
    import fitz
    adv = collections.defaultdict(list)
    faces = collections.Counter()
    for pg in fitz.open(pdf):
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                for s in ln["spans"]:
                    sc = s["chars"]
                    faces[s["font"]] += len(sc)
                    for i, c in enumerate(sc[:-1]):
                        if not (A.is_cjk(c["c"]) and A.is_cjk(sc[i + 1]["c"])):
                            continue
                        a = (sc[i + 1]["bbox"][0] - c["bbox"][0]) / s["size"]
                        if 0.2 < a < 1.3:
                            adv[(s["font"], c["c"])].append(a)
    face = faces.most_common(1)[0][0] if faces else "?"
    med = {}
    for (f, ch), v in adv.items():
        if f != face or len(v) < 3:
            continue
        med[ch] = sorted(v)[len(v) // 2]
    return face, med


def main():
    os.makedirs(OUT, exist_ok=True)
    keep = "--keep" in sys.argv
    upm, widths = A.design_table("MS PGothic")
    # kana the design table ships narrow -- where a full-em layout is visible
    probes = [ch for ch in "のにはるまいてしとこ"
              if widths.get(str(ord(ch))) and widths[str(ord(ch))] / upm < 0.95]
    print("design (MS PGothic): " +
          " ".join("%s%.3f" % (ch, widths[str(ord(ch))] / upm) for ch in probes))
    print("%-16s %-12s %-12s %s" % ("docGrid", "face drawn", "kanji em", "kana advances"))
    for grid, font in ARMS:
        tag = "".join(c for c in (grid + font) if c.isalnum())
        docx = os.path.join(OUT, "cbgrid_%s.docx" % tag)
        if not (keep and os.path.exists(docx)):
            G.build(docx, jc="left", compat="15", grid=grid, pitch="360",
                    font=font, sz="21", ind0=0.0, ind1=24.0, step=6.0,
                    base="1", cpunct="1")
        face, med = measure(export(docx))
        kanji = [v for ch, v in med.items() if 0x4E00 <= ord(ch) <= 0x9FFF]
        kj = sorted(kanji)[len(kanji) // 2] if kanji else float("nan")
        shown = " ".join("%s%.3f" % (ch, med[ch]) for ch in probes if ch in med)
        print("%-16s %-12s %-12.3f %s" % (grid, face, kj, shown or "(no samples)"))


if __name__ == "__main__":
    main()
