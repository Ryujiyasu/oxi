# -*- coding: utf-8 -*-
"""Where is the page-bottom flip for a 2-line CJK paragraph in a typed grid?

reports__11b5f886 is the last blind-set residual: its 2-line paragraph
「居住系サービス…」 has line0 sitting at cursor 723.30 with the body bottom at
739.85 and a 16.35pt grid line, so line0 clears the bottom by 0.2pt -- Oxi keeps
it and splits 1+1, Word moves the whole paragraph to the next page.  Neither of
the rules already derived predicts that:

  * widowControl is OFF (Normal declares w:val="0"), so the widow look-ahead
    never arms -- and _pb_orphan_gen.py shows Word's orphan protection under
    widowControl=0 matches Oxi on 7 Latin arms.
  * S605/S748 push line0 only when its FULL cell overflows; here it fits.
  * S748's centered-box flip, (pitch + natural)/2 = (16.35 + 10.375)/2 = 13.36,
    would keep line0 with 16.55pt of room.

So sweep the cursor across the boundary in this document's own regime (grid
linePitch 327, HG丸ｺﾞｼｯｸM-PRO 8pt, margins 2041) and read Word's flip directly.
An `exact` spacer paragraph moves the cursor in sub-point steps without taking a
grid slot.

  python _pb_cjk2line_gen.py gen
  python _pb_cjk2line_gen.py pdf      # Word truth
  python _pb_cjk2line_gen.py oxi      # Oxi, same arms
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_cjk2line")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = "HG丸ｺﾞｼｯｸM-PRO"
SZ_HP = 16                 # 8pt, as in the specimen
PITCH = 327                # twips = 16.35pt
# kojin (the KEEP counterexample) uses a linesAndChars grid; this probe used
# type=lines. OXI_PB_GRID switches it to test the regime as the discriminator.
GRID_TYPE = os.environ.get("OXI_PB_GRID", "lines")
CHAR_SPACE = int(os.environ.get("OXI_PB_CHARSPACE", "0"))
FILLERS = 36               # grid lines before the spacer
EMPTIES = 7                # of those, empty paragraphs immediately before
# spacer heights in twips; 20tw = 1pt. The flip is expected somewhere in here.
# refined sweep: 0.25pt (5tw) steps across both edges of the split window
# third sweep: around the specimen's own position. cursor = 710.45 +
# (tw-70)/20, so tw=327 reproduces its 723.30 (room 16.55, line box 16.35
# -> line0 clears the bottom by 0.20 and Word still moves the paragraph).
# fourth sweep: the LAST line's natural slack. slack = 2.675 - (tw-70)/20,
# so tw 110..132 walks it from 0.675 down to -0.425 in 0.1pt steps. kojin
# pi=52 (Word KEEPS) sits at slack 0.35 = tw 116.5; this probe's earlier
# arms (slack 0.55..2.675) all BREAK, so the flip is in here.
SPACERS = list(range(110, 134, 4))
LINE1 = ("居住系サービス　共同生活援助（グループホーム）・共同生活介護（ケアホーム）　"
         "利用者数　23年度実績　5,921　24年度　見込み　6,374　実績　6,635　"
         "25年度　見込み　6,907　実績　7,321　26年度　見込み　7,441")


def docx():
    return os.path.join(OUT, "cjk2line.docx")


def para(text, spacing="", sz=SZ_HP):
    return ('<w:p><w:pPr>%s</w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (spacing, FACE, FACE, FACE, sz, sz, text))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, tw in enumerate(SPACERS):
        body.append(para("A%02dZ" % ai,
                         "<w:pageBreakBefore/>" if ai else ""))
        for k in range(FILLERS - EMPTIES):
            body.append(para("うめ%02d-%02d" % (ai, k)))
        # S-specimen shape: the test paragraph is preceded by a run of EMPTY
        # paragraphs (reports__11b5f886 has ~7 grid lines of them, 608.9 -> 723.3).
        # Text fillers alone did not reproduce Word's whole-paragraph push.
        for _ in range(EMPTIES):
            body.append(para(""))
        # exact-height spacer: moves the cursor without claiming a grid slot
        body.append(para("s",
                         '<w:spacing w:before="0" w:after="0" w:line="%d"'
                         ' w:lineRule="exact"/>' % tw))
        body.append(para("T%02d %s" % (ai, LINE1)))
        body.append(para("E%02dZ" % ai))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
           '<w:pgMar w:top="2041" w:right="851" w:bottom="2041" w:left="851" '
           'w:header="851" w:footer="992" w:gutter="0"/>'
           '<w:docGrid w:type="%s" w:linePitch="%d" w:charSpace="%d"/>'
           "</w:sectPr></w:body></w:document>" % (GRID_TYPE, PITCH, CHAR_SPACE))
    # widowControl OFF in the default style, exactly like the specimen
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
    print("wrote", docx(), len(SPACERS), "arms; pitch %.2fpt" % (PITCH / 20.0))


def report(per, who):
    print("== %s ==" % who)
    print("%-8s %9s %-18s %s" % ("spacer", "pt", "line0/line1 page", "verdict"))
    for ai, tw in enumerate(SPACERS):
        g = per.get(ai)
        if not g or "l0" not in g:
            print("%-8s %9.2f MISSING" % ("sp%d" % tw, tw / 20.0))
            continue
        l0, l1 = g.get("l0"), g.get("l1")
        v = "split 1+1" if (l1 is not None and l1 != l0) else "both on p%s" % l0
        print("%-8s %9.2f %-18s %s" % ("sp%d" % tw, tw / 20.0, "%s/%s" % (l0, l1), v))


def _pages(textpages):
    per = {}
    for pi, txt in enumerate(textpages):
        for m in re.finditer(r"T(\d\d)", txt):
            per.setdefault(int(m.group(1)), {}).setdefault("l0", pi + 1)
        # line 1 = the tail of the long paragraph; use its last token
        for m in re.finditer(r"7,441", txt):
            pass
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
    per = {}
    for pi in range(doc.page_count):
        t = doc[pi].get_text()
        for m in re.finditer(r"T(\d\d)", t):
            per.setdefault(int(m.group(1)), {}).setdefault("l0", pi + 1)
        if "7,441" in t:
            # attribute the tail to the arm whose T-marker is on this or the
            # previous page
            for ai in per:
                if per[ai].get("l0") in (pi, pi + 1) and "l1" not in per[ai]:
                    per[ai]["l1"] = pi + 1
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "cjk2line_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "c2"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    per = {}
    for pi, pg in enumerate(pages):
        t = "".join(e.get("text") or "" for e in pg["elements"] if e["type"] == "text")
        for m in re.finditer(r"T(\d\d)", t):
            per.setdefault(int(m.group(1)), {}).setdefault("l0", pi + 1)
        if "7,441" in t:
            for ai in per:
                if per[ai].get("l0") in (pi, pi + 1) and "l1" not in per[ai]:
                    per[ai]["l1"] = pi + 1
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
