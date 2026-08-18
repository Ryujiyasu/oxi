# -*- coding: utf-8 -*-
"""How much does a 約物 give up when a cell line is short of room?

The corpus scan could not answer this: indents, merged cells, undrawn borders and
per-run `w:spacing` all move the same numbers, and 11% of lines with room to spare
came back "compressed". So control it. One string, one face, one size, and only the
cell width moves -- a twip at a time -- so the line's shortfall walks from "-2em of
slack" to "+2em of overflow" and the response can be read straight off.

    demand   = natural width of the line - the cell's content width
    supplied = natural width - what Word drew

Arms vary what the line has to give: a closing bracket at the end (34140b's case),
nothing at all (does plain text compress, or does it just wrap?), a bracket in the
middle, 読点/句点, and two candidates competing.

    python _cw_yaku_probe.py            # generate, export through Word, measure
    python _cw_yaku_probe.py --keep     # reuse the export
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import _cw_law as L  # noqa: E402

MARK = "甲"
# Every document in the corpus that shows in-cell compression declares this.
CPUNCT = os.environ.get("CPUNCT", "compressPunctuation")
# 34140b's compressing cell is CENTRED and its document declares no
# compatibilityMode at all -- both are axes, not decoration.
JC = os.environ.get("JC", "left")
COMPAT = os.environ.get("COMPAT", "15")
EM = L.SZ / 2.0                       # 10.5pt, and ＭＳ 明朝 is fullwidth
ARMS = {
    "Y1": ("甲　年度）", "closing bracket last -- 34140b's line"),
    "Y2": ("甲亜亜亜亜", "no 約物 at all -- the control"),
    "Y3": ("甲）年度亜", "closing bracket in the middle"),
    "Y4": ("甲亜亜亜、", "読点 last"),
    "Y5": ("甲亜亜亜。", "句点 last"),
    "Y6": ("甲）亜亜）", "two closing brackets competing"),
    "Y7": ("甲（亜亜亜", "an OPENING bracket -- does it give too?"),
    "Y8": ("甲亜亜亜　", "a trailing ideographic space"),
    "Y9": ("甲亜　亜亜", "an ideographic space MID-line -- does it shrink on a justified line?"),
}
# 5 characters at 10.5pt = 52.5pt of text. Sweep the content area from 42 to 63pt,
# i.e. demand from +10.5 (one em over) to -10.5 (one em of slack).
W0, W1 = 42.0 + 10.8, 63.0 + 10.8      # + both cell margins, in points
L.WINDOWS = [(W0 * 20, W1 * 20, 1.0)]


def build(out, text):
    L.NCHAR = len(text)
    src = L.MARK
    L.MARK = text[0]
    try:
        original = L.FILL
        # build() writes MARK + FILL*(NCHAR-1); patch in the exact string instead
        import zipfile  # noqa: F401
        L.FILL = None
        _build_exact(out, text)
    finally:
        L.MARK = src
        L.FILL = original


def _build_exact(out, text):
    """L.build with the swept text replaced by this arm's exact string."""
    import re
    import zipfile
    tmp = out + ".tmp"
    L.MARK, keep = text[0], L.MARK
    L.FILL = text[1] if len(text) > 1 else "亜"
    L.NCHAR = len(text)
    cfg = dict(L.ARMS["A"])
    cfg["cpunct"] = CPUNCT
    cfg["jc"] = JC
    cfg["compat"] = COMPAT
    L.build(tmp, **cfg)
    L.MARK = keep
    z = zipfile.ZipFile(tmp)
    parts = {n: z.read(n) for n in z.namelist()}
    z.close()
    os.remove(tmp)
    doc = parts["word/document.xml"].decode("utf-8")
    wrong = text[0] + L.FILL * (len(text) - 1)
    doc = doc.replace("<w:t>%s</w:t>" % wrong, "<w:t>%s</w:t>" % text)
    parts["word/document.xml"] = doc.encode("utf-8")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as w:
        for n, d in parts.items():
            w.writestr(n, d)


def measure(pdf, text):
    """Per swept width: did the line hold, and what did each character take?"""
    import fitz
    cells, cur = [], None
    for pg in fitz.open(pdf):
        rows = []
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                cs = [(c["c"], c["bbox"][0]) for sp in ln["spans"] for c in sp["chars"]]
                if cs:
                    rows.append((round(ln["bbox"][1], 1), cs))
        rows.sort()
        for _, cs in rows:
            if cs[0][0] == text[0]:
                cur = []
                cells.append(cur)
            if cur is not None:
                cur.append(cs)
    return cells


def main():
    os.makedirs(L.OUT, exist_ok=True)
    keep = "--keep" in sys.argv
    want = [a for a in sys.argv[1:] if a in ARMS] or list(ARMS)
    ws = L.widths()
    print(f"characterSpacingControl={CPUNCT or '(absent)'}  jc={JC}  compat={COMPAT}")
    print(f"sweep {len(ws)} widths, content area {W0 - 10.8:.1f}..{W1 - 10.8:.1f}pt, "
          f"natural line = 5 x {EM} = {5 * EM}pt")
    for name in want:
        text, note = ARMS[name]
        docx = os.path.join(L.OUT, f"cwy_{name}.docx")
        if not (keep and os.path.exists(docx)):
            _build_exact(docx, text)
        cells = measure(L.export(docx), text)
        if len(cells) != len(ws):
            print(f"{name}: {len(cells)} cells vs {len(ws)} widths -- skipped")
            continue
        print(f"\n=== {name} {text!r} -- {note}")
        show_curve = "--curve" in sys.argv
        if show_curve:
            print(f"    {'demand':>8}{'held':>7}{'supplied':>10}   who gave up")
        prev = None
        gave = {}
        err = []
        hold_max = None
        for w, lines in zip(ws, cells):
            inner = w / 20.0 - 10.8
            first = lines[0]
            held = len(first) >= len(text)
            demand = round(len(text) * EM - inner, 2)
            if held:
                nxt = first[len(text)][1] if len(first) > len(text) else None
                if nxt is None:
                    continue
                adv = [round(first[i + 1][1] - first[i][1], 2) for i in range(len(text))]
                supplied = round(len(text) * EM - sum(adv), 2)
                who = " ".join(f"{text[i]}{EM - adv[i]:+.2f}" for i in range(len(text))
                               if abs(EM - adv[i]) > 0.15) or "-"
            else:
                supplied, who = None, f"WRAPPED after {len(first)}"
            # one row per half-point of demand: the response is continuous, and
            # printing every twip buries the shape in 421 near-identical lines.
            if held:
                for i in range(len(text)):
                    g = EM - adv[i]
                    if g > 0.15:
                        gave[text[i]] = max(gave.get(text[i], 0.0), g)
                err.append(supplied - demand)
                if hold_max is None or demand > hold_max:
                    hold_max = demand
            key = (held, who, round(demand * 2) / 2)
            if key != prev:
                if show_curve:
                    s = f"{supplied:>10.2f}" if supplied is not None else f"{'-':>10}"
                    print(f"    {demand:>+8.2f}{'yes' if held else 'no':>7}{s}   {who}")
                prev = key
        if gave:
            print("    gives up (max, and as a fraction of the em): "
                  + "  ".join(f"{c}{g:.2f}={g / EM:.3f}em" for c, g in
                              sorted(gave.items(), key=lambda kv: -kv[1])))
        else:
            print("    gives up nothing -- the line wraps instead")
        if err:
            err.sort()
            print(f"    supplied - demand over {len(err)} held widths: "
                  f"median {err[len(err) // 2]:+.3f}, range {err[0]:+.3f}..{err[-1]:+.3f}")
            print(f"    holds up to demand {hold_max:+.2f}pt = {hold_max / EM:.3f} em")


if __name__ == "__main__":
    main()
