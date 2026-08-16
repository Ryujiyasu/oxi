# -*- coding: utf-8 -*-
"""Does a typed grid's LAST line get page-bottom leniency? Sweep phase x slack.

_pb_cjk2line_gen.py showed Word splitting a 2-line CJK paragraph 1+1 in all its
arms where Oxi's natural-height leniency keeps both lines -- i.e. Word used the
FULL grid box on the last line.  Applying that unconditionally (OXI_S1152=1)
takes the probe 2/6 -> 6/6 but drops Phase 1 from 95 to 90 (34140, db9ca,
ohnochingin, roudoujoken, tokyoshugyo), so the leniency IS real on those pages
and the probe's regime is narrower than "any typed-grid last line".

The probe and those documents differ in two things at once, and the earlier
sweeps could not separate them because moving the spacer moves BOTH:

    slot phase   -- the exact-height spacer leaves the cursor off-slot
    slack        -- how far the natural height clears the content bottom

So sweep them independently.  The spacer sets the phase (and, incidentally, the
slack); the section's BOTTOM MARGIN sets the slack alone.  Each arm is its own
section so it carries its own bottom margin.

    python _pb_lastline_gen.py gen
    python _pb_lastline_gen.py pdf      # Word truth
    python _pb_lastline_gen.py oxi      # Oxi, same arms
    python _pb_lastline_gen.py oxi OXI_S1152=1
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_lastline")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = "HG丸ｺﾞｼｯｸM-PRO"
SZ_HP = 16                 # 8pt
PITCH = 327                # twips = 16.35pt
PGH = 16838
TOP = 2041                 # twips
FILLERS = 36
EMPTIES = 7
NAT = 10.375               # natural line height at 8pt in this face (BR_DUMP)
BOX = PITCH / 20.0         # 16.35pt grid cell

# phase: spacer height in twips. 0 keeps the cursor on the slot the fillers
# left it on; PITCH would land on the next slot, so sweep [0, PITCH).
PHASES = [0, 82, 164, 245]
LINE1 = ("居住系サービス　共同生活援助（グループホーム）・共同生活介護（ケアホーム）　"
         "利用者数　23年度実績　5,921　24年度　見込み　6,374　実績　6,635　"
         "25年度　見込み　6,907　実績　7,321　26年度　見込み　7,441")
# OXI_PB_LINES=3 makes the test paragraph wrap to THREE lines, so the line the
# sweep puts at the page bottom (index 1) becomes a NON-LAST line at exactly the
# same cursor and slack. That is the discriminator between "the centered box is
# the rule for every typed-grid line" and S693's "non-last hairline -> full box,
# non-last comfortable -> ink leniency": in the fine window nat_over is -2.5 to
# -3.2, i.e. comfortable, so S693 predicts KEEP throughout while the centered box
# predicts the same 2.9875 flip as the 2-line sweep.
TAIL = ("　27年度　見込み　7,560　実績　7,684　28年度　見込み　7,802　実績　7,915"
        "　29年度　見込み　8,031　実績　8,142　30年度　見込み　8,260")
if os.environ.get("OXI_PB_LINES") == "3":
    LINE1 = LINE1 + TAIL


def _bottom_for(phase, natslack):
    """Bottom margin (tw) that gives the last line this natural slack.

    natslack = bot - c_line1 - NAT, and bot moves 1pt per 20tw of margin, so
    the margin is solved directly instead of swept blindly -- that is what
    lets phase and slack move independently.
    """
    page_top = TOP / 20.0
    c1 = page_top + (FILLERS + 2) * BOX + phase / 20.0
    return int(round(PGH - TOP - 20.0 * (c1 + NAT + natslack - page_top)))


# slack ladder in 0.5pt steps: the first sweep bracketed the flip between
# natslack 2.125 (Word splits) and 4.125 (Word keeps) at phase 0.
# OXI_PB_FINE=1 walks the bracket the coarse ladder left (2.5 .. 3.2 by 0.1),
# where the centered box (BOX+NAT)/2 predicts the flip at natslack 2.9875.
if os.environ.get("OXI_PB_FINE"):
    SLACKS = [round(3.2 - 0.1 * k, 3) for k in range(8)]
else:
    SLACKS = [round(6.0 - 0.5 * k, 3) for k in range(14)]
ARMS = [(p, _bottom_for(p, s)) for p in PHASES for s in SLACKS]


def docx():
    return os.path.join(OUT, "lastline.docx")


def para(text, ppr=""):
    return ('<w:p><w:pPr>%s</w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (ppr, FACE, FACE, FACE, SZ_HP, SZ_HP, text))


def sect(bottom):
    return ('<w:pgSz w:w="11906" w:h="%d" w:code="9"/>'
            '<w:pgMar w:top="%d" w:right="851" w:bottom="%d" w:left="851" '
            'w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:docGrid w:type="lines" w:linePitch="%d" w:charSpace="0"/>'
            % (PGH, TOP, bottom, PITCH))


def geometry(phase, bottom):
    """(cursor of line1, content bottom, natural slack, phase within cell).

    Lines ahead of the test paragraph: the A-marker + (FILLERS-EMPTIES) text
    fillers + EMPTIES empties = FILLERS+1 grid lines. (Counting FILLERS put
    every arm one whole cell off and made the first read of this table look
    like Word was breaking lines whose full box still fit.)
    """
    page_top = TOP / 20.0
    c_line1 = page_top + (FILLERS + 2) * BOX + phase / 20.0
    bot = page_top + (PGH - TOP - bottom) / 20.0
    return c_line1, bot, bot - (c_line1 + NAT), (c_line1 - page_top) % BOX


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (phase, bottom) in enumerate(ARMS):
        body.append(para("A%02dZ" % ai))
        for k in range(FILLERS - EMPTIES):
            body.append(para("うめ%02d-%02d" % (ai, k)))
        for _ in range(EMPTIES):
            body.append(para(""))
        if phase:
            body.append(para("s", '<w:spacing w:before="0" w:after="0"'
                                  ' w:line="%d" w:lineRule="exact"/>' % phase))
        body.append(para("T%02d %s" % (ai, LINE1)))
        # the section break carries THIS arm's bottom margin
        body.append('<w:p><w:pPr><w:sectPr>%s</w:sectPr></w:pPr></w:p>'
                    % sect(bottom))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) + "<w:sectPr>" + sect(ARMS[-1][1]) +
           "</w:sectPr></w:body></w:document>")
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr>'
              '<w:rPr><w:sz w:val="%d"/></w:rPr></w:style>'
              "</w:styles>" % (FACE, FACE, FACE, SZ_HP))
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms")


def report(per, who):
    print("== %s ==" % who)
    print("%-4s %-7s %-8s %-9s %-9s %s"
          % ("arm", "phase", "bottom", "natslack", "boxover", "verdict"))
    for ai, (phase, bottom) in enumerate(ARMS):
        c1, bot, slack, _ = geometry(phase, bottom)
        g = per.get(ai) or {}
        a, l0, l1 = g.get("a"), g.get("l0"), g.get("l1")
        if l0 is None:
            v = "MISSING"
        elif a is not None and l0 > a:
            # line 0 itself did not fit -- the whole paragraph moved, which
            # says nothing about the LAST line's threshold
            v = "move-whole"
        elif l1 is None:
            v = "no-tail"
        elif l1 != l0:
            v = "SPLIT"
        else:
            v = "keep"
        print("%-4d %-7.2f %-8d %-9.3f %-9.3f %s"
              % (ai, phase / 20.0, bottom, slack, c1 + BOX - bot, v))


def _collect(pagetexts):
    per = {}
    for pi, t in enumerate(pagetexts):
        for m in re.finditer(r"A(\d\d)Z", t):
            per.setdefault(int(m.group(1)), {}).setdefault("a", pi + 1)
        for m in re.finditer(r"T(\d\d)", t):
            per.setdefault(int(m.group(1)), {}).setdefault("l0", pi + 1)
        if "7,441" in t:
            for ai in per:
                if per[ai].get("l0") in (pi, pi + 1) and "l1" not in per[ai]:
                    per[ai]["l1"] = pi + 1
    return per


def pdf():
    import fitz
    import win32com.client as w
    out = docx().replace(".docx", ".pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(out, 17)
    finally:
        d.Close(False)
        app.Quit()
    doc = fitz.open(out)
    report(_collect([doc[i].get_text() for i in range(doc.page_count)]), "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "lastline_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "ll"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    texts = ["".join(e.get("text") or "" for e in pg["elements"] if e["type"] == "text")
             for pg in pages]
    report(_collect(texts), "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
