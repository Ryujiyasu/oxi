# -*- coding: utf-8 -*-
"""What is the page-bottom EMPTY paragraph's keep tolerance a function of?

_pb_emptytail_gen.py killed every structural candidate: all five successor
variants (nothing / 1 line / 4 lines / 4 lines with widowControl / another
empty) flip at exactly the same overflow, so the successor plays no part.  What
it left is a tolerance, and one number in that run is suggestive --

    Calibri 11 x1.15   lh 15.4419   natural 13.4278   lh - natural = 2.014
    measured flip window                              over in (1.82, 2.32]

i.e. the empty may hang past the content bottom by exactly the leading the
line-spacing multiplier added.  That has a mechanism: Word's multiple line
spacing puts the extra space BELOW the text, so a test run on the natural
(unmultiplied) height lets the added leading overflow.  It also re-reads the
S736 derivation specimen -- probexempty's recorded flip is (1.8, 2.1], which
brackets 2.014 rather than the 2.5 constant that was fitted to it.

Rival models, and where they disagree:

    a  constant ~2.0pt          flat everywhere
    b  lh - natural             zero at mult 1.0, 6.7 at mult 1.5
    c  proportional to size     4.0 at 22pt regardless of mult
    d  lh - ink height          smaller than b, and font-shape dependent

Each combo below sweeps a shim of EXACT line height at the top of the arm in
0.25pt steps, walking the stack across the content bottom over a window that
always contains 0 (the strict box) and 2.5 (the S736 constant) as well as the
model-b prediction.  The body line count is chosen per combo so that the last
body line always fits.

★RESULT (2026-08-13): model b, and only model b.  Every one of the 14 combos is
monotone and its flip window contains lh - natural:

    Calibri 11 x1.000  (-0.23, 0.02] ∋ 0.00     Calibri 22 x1.000 (-0.23, 0.02] ∋ 0.00
    Calibri 11 x1.079  ( 0.98, 1.23] ∋ 1.06     Calibri 22 x1.150 ( 4.02, 4.27] ∋ 4.03
    Calibri 11 x1.150  ( 2.01, 2.26] ∋ 2.01     Arial   12 x1.150 ( 2.01, 2.26] ∋ 2.07
    Calibri 11 x1.500  ( 6.51, 6.76] ∋ 6.71     TNR     12 x1.500 ( 6.77, 7.02] ∋ 6.90

A constant is refuted by 6 of the 8 auto arms, k*size by the x1.000 arms and by
Times.  atLeast 18pt / exact 18pt / exact 10pt all flip at (0.00, 0.25] = the
full box: the multiplier is the ONLY source of hangable leading.  And the
empty's own w:after is not part of the test (after=200 leaves the flip alone).
Shipped as S1113.

  python _pb_emptytol_gen.py gen
  python _pb_emptytol_gen.py read          # Word COM truth (caches it)
  python _pb_emptytol_gen.py oxi OXI_S1113=1   # Oxi per-arm agreement vs cache
"""
import math
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_emptytail")
DOCX = os.path.join(OUT, "emptytol.docx")

sys.path.insert(0, HERE)
from _pb_emptyrun_gen import natural  # noqa: E402
from _pb_pxgrid_gen import CT, DRELS, NS, RELS, STYLES  # noqa: E402

TOP_TW, BOTTOM_TW = 1440, 8200         # 72.0pt / 410.0pt -> content bottom 431.90
PAGE_H = 16838 / 20.0
TOP = TOP_TW / 20.0
CBOT = PAGE_H - BOTTOM_TW / 20.0
STEP = 0.25                            # shim resolution, pt
H_MIN = 1.0                            # smallest shim we are willing to emit
LEAD_IN = 0.5                          # how far below over=0 the sweep starts

# (font, half-point size, w:line, w:lineRule, w:after on the empty only)
COMBOS = [
    ("Calibri", 22, 240, "auto", 0),
    ("Calibri", 22, 259, "auto", 0),
    ("Calibri", 22, 276, "auto", 0),
    ("Calibri", 22, 360, "auto", 0),
    ("Calibri", 44, 240, "auto", 0),
    ("Calibri", 44, 276, "auto", 0),
    ("Arial", 24, 276, "auto", 0),
    ("Times New Roman", 24, 360, "auto", 0),
    # does the multiplier's leading generalise to the other two line rules?
    ("Calibri", 22, 360, "atLeast", 0),     # 18pt requested, natural 13.43
    ("Calibri", 22, 240, "atLeast", 0),     # 12pt requested -> natural wins
    ("Calibri", 22, 360, "exact", 0),       # 18pt box, no natural involved
    ("Calibri", 22, 200, "exact", 0),       # 10pt box, below natural
    # and does the empty's own space-after have to fit as well?
    ("Calibri", 22, 276, "auto", 200),      # 10pt after on the empty only
    ("Calibri", 22, 240, "auto", 200),
]


def line_height(h1, line, rule):
    if rule == "auto":
        return h1 * (line / 240.0)
    if rule == "exact":
        return line / 20.0
    return max(h1, line / 20.0)             # atLeast


def plan():
    """Per combo: line height, natural height, body count, shim list."""
    nat = natural()
    out = []
    for font, sz, line, rule, after in COMBOS:
        h1 = nat[font] * (sz / 2.0)
        lh = line_height(h1, line, rule)
        width = max(lh - h1, 2.5, after / 20.0) + 2.0
        nb = int((CBOT - TOP - H_MIN - LEAD_IN) / lh) - 1        # body lines
        h0 = CBOT - TOP - (nb + 1) * lh - LEAD_IN                # over(h0) = -LEAD_IN
        # the shim is authored in whole twips, so quantise here and let `over`
        # be computed from the value Word actually gets
        shims = [round(h0 * 20 + i * STEP * 20) / 20.0
                 for i in range(int(math.ceil(width / STEP)) + 1)]
        out.append(dict(font=font, sz=sz, line=line, rule=rule, after=after,
                        h1=h1, lh=lh, nb=nb, shims=shims))
    return out


def arms():
    return [(c, h) for c in plan() for h in c["shims"]]


def rpr(font, sz):
    return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            % (font, font, font, sz, sz))


def ppr(c, pbb=False, shim=None, after=0):
    p = []
    if pbb:
        p.append("<w:pageBreakBefore/>")
    p.append('<w:widowControl w:val="0"/>')
    if shim is None:
        p.append('<w:spacing w:before="0" w:after="%d" w:line="%d" w:lineRule="%s"/>'
                 % (after, c["line"], c["rule"]))
    else:
        p.append('<w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/>' % shim)
    p.append(rpr(c["font"], c["sz"]))
    return "<w:pPr>%s</w:pPr>" % "".join(p)


def gen():
    os.makedirs(OUT, exist_ok=True)
    body, total = [], 0
    for ai, (c, h) in enumerate(arms()):
        f, sz, nb = c["font"], c["sz"], c["nb"]
        body.append("<w:p>%s</w:p>" % ppr(c, pbb=True, shim=int(round(h * 20))))
        for k in range(nb):
            tag = "T%03d" % ai if k == 0 else ("B%03d" % ai if k == nb - 1 else "x")
            body.append("<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>"
                        % (ppr(c), rpr(f, sz), tag))
        body.append("<w:p>%s</w:p>" % ppr(c, after=c["after"]))   # the empty under test
        body.append("<w:p>%s<w:r>%s<w:t>S%03d</w:t></w:r></w:p>"
                    % (ppr(c), rpr(f, sz), ai))
        total += nb + 3
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="%d" w:right="1440" w:bottom="%d" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
           % (TOP_TW, BOTTOM_TW))
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(arms()), "arms /", total, "paragraphs")


def indices():
    """1-based index of (first body, last body, empty) per arm."""
    out, base = [], 0
    for c, _h in arms():
        nb = c["nb"]
        out.append((base + 2, base + 1 + nb, base + 2 + nb))
        base += nb + 3
    return out


def probe_word():
    import win32com.client as w
    wanted = {}
    for ai, (first, last, emp) in enumerate(indices()):
        wanted[first] = (ai, "first")
        wanted[last] = (ai, "last")
        wanted.setdefault(emp, (ai, "empty"))
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.ScreenUpdating = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
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
            rows.setdefault(ai, {})[slot] = (rng.Text.rstrip("\r\x07"),
                                             c.Information(3), round(c.Information(6), 2))
        for ai in range(len(arms())):
            r = rows.get(ai, {})
            if r.get("first", ("",))[0] != "T%03d" % ai or r.get("last", ("",))[0] != "B%03d" % ai:
                bad.append(ai)
    finally:
        d.Close(False)
        app.Quit()
    if bad:
        raise SystemExit("paragraph index drifted on arms %s" % bad[:8])
    return rows


WORD_CACHE = os.path.join(OUT, "emptytol_word.json")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")


def flip_window(seq):
    """(last KEEP overflow, first PUSH overflow, monotone?) for one combo."""
    keeps = [o for o, k in seq if k]
    pushes = [o for o, k in seq if not k]
    return (max(keeps) if keeps else None,
            min(pushes) if pushes else None,
            not keeps or not pushes or max(keeps) < min(pushes))


def oxi(envs=""):
    """Oxi's own KEEP/PUSH per arm, from the layout dump, vs the Word cache."""
    import json
    import subprocess
    import tempfile
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "emptytol_oxi.json")
    subprocess.run([GDI, DOCX, os.path.join(tempfile.gettempdir(), "etol"),
                    "--dump-layout=" + out], check=True, env=env,
                   capture_output=True)
    page_of = {}
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        for el in pg["elements"]:
            page_of.setdefault(el.get("para_idx"), pg["page"])
    y_of = {}
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        for el in pg["elements"]:
            y_of.setdefault(el.get("para_idx"), el["y"])
    word = json.load(open(WORD_CACHE, encoding="utf-8")) if os.path.exists(WORD_CACHE) else {}
    print("env: %s\n" % (envs or "(default)"))
    print("%-14s %4s %5s %-8s %4s %8s | %8s %8s %s"
          % ("font", "sz", "line", "rule", "aft", "b:lh-nat",
             "arms=W", "drift", "first disagreement"))
    per, base = {}, 0
    for _ai, (c, h) in enumerate(arms()):
        nb = c["nb"]
        first, last, emp = base + 1, base + nb, base + 1 + nb    # 0-based
        base += nb + 3
        if last not in page_of or emp not in page_of:
            continue
        over = TOP + h + (nb + 1) * c["lh"] - CBOT
        # span first->empty cancels text_y_off, so this is pure cursor drift --
        # but only while both sit on the same page
        drift = ((y_of[emp] - y_of[first]) - c["nb"] * c["lh"]
                 if page_of[emp] == page_of[first] else None)
        per.setdefault(id(c), (c, []))[1].append(
            (over, page_of[emp] == page_of[last], drift))
    hit = tot = 0
    for c, seq in per.values():
        w = word.get("%s|%d|%d|%s|%d" % (c["font"], c["sz"], c["line"], c["rule"], c["after"]))
        wk = dict(w["arms"]) if w else {}
        same = [round(o, 2) in wk and wk[round(o, 2)] == k for o, k, _dy in seq]
        first_bad = next((i for i, s in enumerate(same) if not s), None)
        hit += sum(same)
        tot += len(same)
        note = "-" if first_bad is None else (
            "over %+.2f: Word %s / Oxi %s"
            % (seq[first_bad][0], "K" if wk.get(round(seq[first_bad][0], 2)) else "P",
               "K" if seq[first_bad][1] else "P"))
        drifts = [dy for _o, _k, dy in seq if dy is not None]
        print("%-14s %4.1f %5d %-8s %4d %8.2f | %4d/%-3d %8s %s"
              % (c["font"], c["sz"] / 2.0, c["line"], c["rule"], c["after"],
                 c["lh"] - c["h1"], sum(same), len(same),
                 "%+.3f" % max(drifts) if drifts else "n/a", note))
    print("\n%d / %d arms match Word" % (hit, tot))


def read():
    import json
    rows = probe_word()
    print("content = [%.2f, %.2f]   step %.2fpt\n" % (TOP, CBOT, STEP))
    per = {}
    for ai, (c, h) in enumerate(arms()):
        r = rows.get(ai)
        if not r or "empty" not in r:
            continue
        over = TOP + h + (c["nb"] + 1) * c["lh"] - CBOT
        keep = r["empty"][1] == r["last"][1]
        per.setdefault(id(c), (c, []))[1].append((h, over, keep, r["first"][2]))
    print("%-14s %4s %5s %-8s %4s %8s %8s | %-16s %6s %8s"
          % ("font", "sz", "line", "rule", "aft", "lh", "natural",
             "flip window", "b:lh-nat", "verdict"))
    for c, seq in per.values():
        keeps = [o for _h, o, k, _y in seq if k]
        pushes = [o for _h, o, k, _y in seq if not k]
        mono = not keeps or not pushes or max(keeps) < min(pushes)
        lo = max(keeps) if keeps else None
        hi = min(pushes) if pushes else None
        b = c["lh"] - c["h1"]
        ok = (lo is None or lo <= b + 1e-6) and (hi is None or b < hi + 1e-6)
        print("%-14s %4.1f %5d %-8s %4d %8.4f %8.4f | (%6s,%6s] %6.2f %8s"
              % (c["font"], c["sz"] / 2.0, c["line"], c["rule"], c["after"],
                 c["lh"], c["h1"],
                 "%.2f" % lo if lo is not None else "-",
                 "%.2f" % hi if hi is not None else "-", b,
                 ("b-FITS" if ok else "b-FAILS") + ("" if mono else "/INTERLEAVED")))
    print("\nper-arm detail (box overflow / verdict), one line per combo:")
    cache = {}
    for c, seq in per.values():
        lo, hi, _mono = flip_window([(o, k) for _h, o, k, _y in seq])
        cache["%s|%d|%d|%s|%d" % (c["font"], c["sz"], c["line"], c["rule"], c["after"])] = \
            {"lo": lo, "hi": hi, "arms": [[round(o, 2), k] for _h, o, k, _y in seq]}
        print("  %-14s %4.1f %5d %-8s aft%-4d %s"
              % (c["font"], c["sz"] / 2.0, c["line"], c["rule"], c["after"],
                 " ".join("%+.2f%s" % (o, "K" if k else "P") for _h, o, k, _y in seq)))
    json.dump(cache, open(WORD_CACHE, "w", encoding="utf-8"), indent=1)
    print("\nWord truth cached to", WORD_CACHE)


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "read": read}[sys.argv[1]]()

