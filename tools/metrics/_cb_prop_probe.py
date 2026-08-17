# -*- coding: utf-8 -*-
"""Does Word take the break-time 約物 credit in a PROPORTIONAL Japanese face?

The corpus census says the credit is worth -520 lines on nedocontract (ＭＳ 明朝)
and +33 on c7b923e5 (ＭＳ Ｐゴシック), and c7b923e5's Word PDF shows every mark at
its natural advance.  The reading is that a proportional face has already spent
the mark's blank half, so Word has nothing to compress and breaks at natural
width, while a monospace face carries the full em Word can take back.

This puts the same 約物-rich sentence through Word in four faces -- two
monospace, two proportional -- at a swept right indent, and asks two questions
per face:

    1. what advance does Word give 、 。 （ ）             (natural or compressed)
    2. does Oxi reproduce Word's lines better with the
       break credit on or off                            (the break itself)

    python _cb_prop_probe.py            # generate, export through Word, report
    python _cb_prop_probe.py --keep     # reuse the PDFs already exported
"""
import collections
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8")

import _cb_budget as B  # noqa: E402
import _cb_gen as G  # noqa: E402

OUT = os.path.join(B.REPO, "pipeline_data", "_cb_prop")
NO_CREDIT = "OXI_S475_SOLO=0,OXI_S475_PAIR=0,OXI_S475_OPEN=0"
FACES = [
    ("mono", "ＭＳ 明朝"),
    ("mono", "ＭＳ ゴシック"),
    ("prop", "ＭＳ Ｐ明朝"),
    ("prop", "ＭＳ Ｐゴシック"),
]
MARKS = "、。（）「」"


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


def mark_advances(pdf):
    """Word's own advance for each mark, as the x step to the next character."""
    import fitz
    adv = collections.defaultdict(list)
    doc = fitz.open(pdf)
    for pg in doc:
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                chars = [c for s in ln["spans"] for c in s["chars"]]
                for i, c in enumerate(chars[:-1]):
                    if c["c"] in MARKS:
                        adv[c["c"]].append(chars[i + 1]["bbox"][0] - c["bbox"][0])
    return {k: sorted(v)[len(v) // 2] for k, v in adv.items() if v}


def main():
    os.makedirs(OUT, exist_ok=True)
    keep = "--keep" in sys.argv
    print("%-6s %-12s %-26s %-11s %-11s %s"
          % ("class", "face", "Word's mark advances", "credit-on", "credit-off", "delta"))
    for cls, face in FACES:
        tag = "".join(ch for ch in face if ch.isalnum()) or cls
        docx = os.path.join(OUT, "cbprop_%s.docx" % tag)
        if not (keep and os.path.exists(docx)):
            G.build(docx, jc="both", compat="15", grid="lines", pitch="360",
                    font=face, sz="21", ind0=0.0, ind1=48.0, step=3.0,
                    base="1", cpunct="1")
        pdf = export(docx)
        adv = mark_advances(pdf)
        shown = " ".join("%s%.2f" % (m, adv[m]) for m in MARKS if m in adv)
        on = B.match_report(docx, "", "on", quiet=True)
        off = B.match_report(docx, NO_CREDIT, "off", quiet=True)
        print("%-6s %-12s %-26s %4d/%-6d %4d/%-6d %+d"
              % (cls, face, shown, on[0], on[1], off[0], off[1], off[0] - on[0]))


if __name__ == "__main__":
    main()
