# -*- coding: utf-8 -*-
"""Does a page-bottom EMPTY paragraph stay because its OWN box fits, or because
its SUCCESSOR does?

Four specimens killed the "tolerance constant" model that S736 implements:

    probexempty               overflow +2.0    Word KEEP
    reference__0042471c       overflow +2.5    Word KEEP
    ukframework  i=400        overflow +0.67   Word PUSH
    correspondence__000f9471  fits by  -0.72   Word PUSH

No monotone threshold keeps +2.5 and pushes a box that fits, so the
discriminator is STRUCTURAL.  The leading candidate is group movement: a
trailing empty is not left alone at the page bottom, so when the paragraph
after it moves to the next page the empty moves with it.  Both PUSH specimens
had a successor that moved; both KEEP specimens were page-terminal.

_pb_emptybot_gen.py could not see this because it read only the EMPTY's page.
This probe reads the empty AND the three paragraphs after it, and separates the
two models by construction:

    empty box bottom  = 411.72 + h      -> fits while h < 20.18   (model BOX)
    successor bottom  = 427.16 + h      -> fits while h <  4.74   (model GROUP)

so the 15.44pt band h in (4.74, 20.18) is where the models disagree: the
empty's own box fits, its successor's does not.  h is a shim paragraph of
EXACT line height at the top of each arm, swept 2.0 .. 24.0pt in 0.5pt steps,
which walks the whole stack across the content bottom without touching any
font metric.

Arms cross that band with five successors:

    none  nothing follows on the page (next arm is pageBreakBefore)
    t1    one 1-line paragraph            fits below the empty until h > 4.74
    t4    a 4-line paragraph, widowControl off   (splits, first line same as t1)
    t4w   a 4-line paragraph, widowControl ON    (never fits: needs 2 lines)
    e1    another empty paragraph          (is the rule about ink?)

t4w is the sharpest arm: its successor always moves for a reason of its own, so
GROUP predicts PUSH across the entire sweep while BOX predicts the same flip at
h = 20.18 as `none`.

★RESULT (2026-08-13): GROUP is FALSIFIED.  All five variants flip at exactly
h in (22.00, 22.50] = box overflow (1.82, 2.32] -- including t4w, whose
successor is on the next page in every single arm.  The successor plays no part
at all, and neither does being page-terminal or having ink.  What remains is a
tolerance, and `_pb_emptytol_gen.py` (run it next) derives it: the fit test uses
the line height BEFORE the lineRule=auto multiplier, so the leading that
multiplier added -- here 15.4419 - 13.4278 = 2.014, right inside the window
above -- is exactly what may hang past the content bottom.

  python _pb_emptytail_gen.py gen  [variant,variant,...]     (default: all five)
  python _pb_emptytail_gen.py read [variant,variant,...]

One document per variant keeps the COM read cheap: Word's Paragraphs.Item(i)
walks from the start of the story, so a 5000-paragraph probe read by index is
quadratic (the first attempt burnt 10 CPU-minutes without finishing).  The
reader below enumerates the collection once and only touches the six
paragraphs per arm whose position it computed in advance.
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_emptytail")

sys.path.insert(0, HERE)
from _pb_emptyrun_gen import natural  # noqa: E402
from _pb_pxgrid_gen import CT, DRELS, NS, RELS, STYLES  # noqa: E402

FONT, SZ, ML = "Calibri", 22, 276      # 11pt x 1.15
NBODY = 21                             # body lines per arm (all fit, always)
TOP_TW, BOTTOM_TW = 1440, 8200         # 72.0pt / 410.0pt -> content bottom 431.90
PAGE_H = 16838 / 20.0                  # 841.90
SHIMS = list(range(40, 481, 10))       # 2.0 .. 24.0pt in 0.5pt steps
VARIANTS = ["none", "t1", "t4", "t4w", "e1"]
ACTIVE = list(VARIANTS)


def docx():
    return os.path.join(OUT, "emptytail_%s.docx" % "-".join(ACTIVE))


def arms():
    return [(v, h) for v in ACTIVE for h in SHIMS]


def indices():
    """1-based paragraph index of (first body, last body, empty) for each arm.

    Arm ai occupies: shim, NBODY body lines, the empty under test, and (unless
    the variant is `none`) one successor.
    """
    out, base = [], 0
    for _ai, (variant, _shim) in enumerate(arms()):
        out.append((base + 2, base + 1 + NBODY, base + 2 + NBODY))
        base += NBODY + 2 + (0 if variant == "none" else 1)
    return out


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            % (FONT, FONT, FONT, SZ, SZ))


def ppr(wc=0, pbb=False, exact=None):
    """pPr in schema order: pageBreakBefore, widowControl, spacing, rPr."""
    p = []
    if pbb:
        p.append("<w:pageBreakBefore/>")
    p.append('<w:widowControl w:val="%d"/>' % wc)
    if exact is None:
        p.append('<w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="auto"/>' % ML)
    else:
        p.append('<w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/>' % exact)
    p.append(rpr())
    return "<w:pPr>%s</w:pPr>" % "".join(p)


def text_p(s, wc=0, nlines=1):
    body = "<w:br/>".join('<w:t xml:space="preserve">%s</w:t>'
                          % (s if i == 0 else "s") for i in range(nlines))
    return "<w:p>%s<w:r>%s%s</w:r></w:p>" % (ppr(wc=wc), rpr(), body)


def successor(ai, variant):
    if variant == "none":
        return ""
    if variant == "e1":
        return "<w:p>%s</w:p>" % ppr()
    return text_p("S%03d" % ai, wc=(1 if variant == "t4w" else 0),
                  nlines=(4 if variant in ("t4", "t4w") else 1))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (variant, shim) in enumerate(arms()):
        body.append("<w:p>%s</w:p>" % ppr(pbb=True, exact=shim))   # the h shim
        for k in range(NBODY):
            tag = "T%03d" % ai if k == 0 else ("B%03d" % ai if k == NBODY - 1 else "x")
            body.append(text_p(tag))
        body.append("<w:p>%s</w:p>" % ppr())                       # the empty under test
        body.append(successor(ai, variant))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="%d" w:right="1440" w:bottom="%d" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
           % (TOP_TW, BOTTOM_TW))
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(arms()), "arms x", NBODY + 2, "paragraphs")


def geometry():
    lh = natural()[FONT] * (SZ / 2.0) * (ML / 240.0)
    top = TOP_TW / 20.0
    cbot = PAGE_H - BOTTOM_TW / 20.0
    return lh, top, cbot


def probe_word():
    """Page + Y of the last body line, the empty, and the 3 paragraphs after it."""
    import win32com.client as w
    wanted = {}
    for ai, (first, last, emp) in enumerate(indices()):
        wanted[first] = (ai, "first")
        wanted[last] = (ai, "last")
        for k in range(4):                       # the empty and what follows it
            wanted.setdefault(emp + k, (ai, "after%d" % k))
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.ScreenUpdating = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    rows, bad = {}, []
    try:
        d.Repaginate()
        i = 0
        for p in d.Paragraphs:
            i += 1
            hit = wanted.get(i)
            if hit is None:
                continue
            ai, slot = hit
            rng = p.Range
            c = d.Range(rng.Start, rng.Start)
            rec = (rng.Text.rstrip("\r\x07"), c.Information(3), round(c.Information(6), 2))
            r = rows.setdefault(ai, {"after": [None] * 4})
            if slot.startswith("after"):
                r["after"][int(slot[5:])] = rec
            else:
                r[slot] = rec
        for ai in range(len(arms())):            # the index arithmetic must hold
            r = rows.get(ai, {})
            if r.get("first", ("",))[0] != "T%03d" % ai or r.get("last", ("",))[0] != "B%03d" % ai:
                bad.append(ai)
    finally:
        d.Close(False)
        app.Quit()
    if bad:
        raise SystemExit("paragraph index drifted on arms %s" % bad[:8])
    return rows


def read():
    lh, top, cbot = geometry()
    rows = probe_word()
    print("lh=%.4f  content=[%.2f, %.2f]  empty bottom=%.2f+h  succ bottom=%.2f+h"
          % (lh, top, cbot, top + (NBODY + 1) * lh, top + (NBODY + 2) * lh))
    print("  model BOX   flips at h=%.2f      model GROUP flips at h=%.2f\n"
          % (cbot - top - (NBODY + 1) * lh, cbot - top - (NBODY + 2) * lh))
    hdr = ("%6s %8s %8s %7s %6s %6s  %-5s  %s"
           % ("h", "y_first", "over", "bodypg", "empty", "succ", "verd", "after (text:page)"))
    flips = {}
    for ai, (variant, shim) in enumerate(arms()):
        r = rows.get(ai)
        if not r or "last" not in r:
            print("%-5s %6.2f  MISSING" % (variant, shim / 20.0))
            continue
        h = shim / 20.0
        over = top + h + (NBODY + 1) * lh - cbot
        after = [a for a in r["after"] if a]
        emp = after[0] if after else None
        keep = emp is not None and emp[1] == r["last"][1]
        if variant not in flips:
            print("\n=== %s ===" % variant)
            print(hdr)
            flips[variant] = []
        flips[variant].append((h, over, keep))
        tail = " ".join("%s:p%d" % (a[0].strip() or "-", a[1]) for a in after[1:])
        print("%6.2f %8.2f %8.2f %7d %6d %6s  %-5s  %s"
              % (h, r["first"][2], over, r["last"][1], emp[1] if emp else -1,
                 ("p%d" % after[1][1]) if len(after) > 1 else "-",
                 "KEEP" if keep else "PUSH", tail))
    print("\n%-6s %-28s %s" % ("var", "flip (last KEEP -> first PUSH)", "monotone"))
    for v in ACTIVE:
        seq = flips.get(v, [])
        keeps = [h for h, _o, k in seq if k]
        pushes = [h for h, _o, k in seq if not k]
        mono = not keeps or not pushes or max(keeps) < min(pushes)
        lo = max(keeps) if keeps else None
        hi = min(pushes) if pushes else None
        print("%-6s %-28s %s"
              % (v, "h in (%s, %s]" % ("%.2f" % lo if lo is not None else "-",
                                       "%.2f" % hi if hi is not None else "-"),
                 "yes" if mono else "NO (interleaved)"))


if __name__ == "__main__":
    if len(sys.argv) > 2:
        ACTIVE = [v for v in sys.argv[2].split(",") if v in VARIANTS]
        assert ACTIVE, "unknown variant %r" % sys.argv[2]
    {"gen": gen, "read": read}[sys.argv[1]]()
