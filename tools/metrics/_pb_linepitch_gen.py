# -*- coding: utf-8 -*-
"""What exactly is Word's single-spaced Latin line height?

Oxi uses (hhea ascent + descent + lineGap) / upm, which for Times New Roman is
(1825 + 443 + 87) / 2048 = 1.14990 em -- 10.349pt at 9pt.  Word's exported PDF
puts the same TOC entry's second line 10.32pt below its first, but that PDF
quantises every y to 1/600 inch (0.12pt), so a single gap cannot separate 10.32
from 10.349: both land on the same 600dpi row.

So measure the pitch the way the quantisation cannot follow: one paragraph long
enough to wrap ~40 times, pitch = (last - first) / (n - 1).  The 0.12pt error
is spent once over 39 gaps, i.e. +-0.003pt per line -- 10x finer than the
difference under test.

Each arm is one page, one font at one size, so the same document answers the
question for every size the specimen's contents block uses (9 / 10 / 11 / 12)
plus two non-Times faces as a control on whether the law is per-font metrics or
a constant.

  python _pb_linepitch_gen.py gen
  python _pb_linepitch_gen.py pdf      # Word truth
  python _pb_linepitch_gen.py oxi      # Oxi, same arms
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_linepitch")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

# (arm, font, half-points)
ARMS = [
    ("tnr8", "Times New Roman", 16),
    ("tnr9", "Times New Roman", 18),
    ("tnr10", "Times New Roman", 20),
    ("tnr11", "Times New Roman", 22),
    ("tnr12", "Times New Roman", 24),
    ("tnr18", "Times New Roman", 36),
    ("tnr105", "Times New Roman", 21),   # half-point size: does Word round it?
    ("arial9", "Arial", 18),
    ("calibri11", "Calibri", 22),
    # Second pass: the law has to hold across the whole face space the EN corpus
    # uses, not just the specimen's Times New Roman -- a per-font exception here
    # would mean the "natural line height" model is wrong rather than incomplete.
    ("cambria11", "Cambria", 22),
    ("georgia10", "Georgia", 20),
    ("verdana9", "Verdana", 18),
    ("segoeui9", "Segoe UI", 18),
    ("tahoma8", "Tahoma", 16),
    ("arialnarrow10", "Arial Narrow", 20),
    ("couriernew10", "Courier New", 20),
    ("trebuchet10", "Trebuchet MS", 20),
    # Third pass: families the corpus names that the metrics table does not hold.
    # Some are Word aliases of a table entry (Helvetica/Courier/Times), some are
    # genuinely absent, and some are not installed at all -- in which case Word
    # substitutes by PANOSE and the arm measures whether Oxi's substitution lands
    # on the same metrics.  A wrong answer here is a wrong line height for every
    # line of every document that names the face.
    ("helvetica10", "Helvetica", 20),
    ("courier10", "Courier", 20),
    ("times10", "Times", 20),
    ("humnst10", "Humnst777 Lt BT", 20),
    ("myriadpro10", "Myriad Pro Light", 20),
    ("futura10", "Futura Bk BT", 20),
    ("grammarsaurus10", "Grammarsaurus", 20),
    ("meiryoui10", "Meiryo UI", 20),
    # calibrate the CJK line-height multiplier against faces already in the
    # table: if Word = natural x 83/64 for these, Meiryo UI's measured
    # 1.65119 / 1.27002 = 1.30013 is the same law within quantisation.
    ("meiryo10", "Meiryo", 20),
    # `line=0 atLeast` (R55: natural height, no grid snap) -- the rule the UD
    # worksheet's body paragraphs carry. Does Word apply the CJK 83/64 there?
    ("min16_at0", "MS Mincho", 32, 'w:before="0" w:after="0" w:line="0" w:lineRule="atLeast"'),
    ("min16_auto", "MS Mincho", 32),
    ("tnr16_at0", "Times New Roman", 32, 'w:before="0" w:after="0" w:line="0" w:lineRule="atLeast"'),
    ("meiryo8", "Meiryo", 16),
    ("meiryo14", "Meiryo", 28),
    ("meiryo20", "Meiryo", 40),
    ("msmincho14", "MS Mincho", 28),
    ("msmincho10", "MS Mincho", 20),
    ("msgothic10", "MS Gothic", 20),
    # Fourth pass: the S1140 batch -- families installed here and named by the
    # corpus that had no table entry. Each new entry is only trustworthy if
    # Word agrees with it on this probe.
    ("segoeuiemoji", "Segoe UI Emoji", 20),
    ("inkfree", "Ink Free", 20),
    ("franklingothi", "Franklin Gothic Book", 20),
    ("wingdings", "Wingdings", 20),
    ("sylfaen", "Sylfaen", 20),
    ("lucidasansuni", "Lucida Sans Unicode", 20),
    ("jokerman", "Jokerman", 20),
    ("impact", "Impact", 20),
    ("erasbolditc", "Eras Bold ITC", 20),
    ("broadway", "Broadway", 20),
    ("baskervilleol", "Baskerville Old Face", 20),
    ("arialroundedm", "Arial Rounded MT Bold", 20),
    ("lucidacallig", "Lucida Calligraphy", 20),
    # S1142 batch: .otf / cloud faces, plus the CJK families the JP corpus
    # names. The CJK arms use Latin text on purpose -- Word applies the
    # 83/64 inflation to a CJK face even then (UD NK-R 1.50379).
    ("montserrat", "Montserrat", 20),
    ("merriweather", "Merriweather", 20),
    ("nunito", "Nunito", 20),
    ("roboto", "Roboto", 20),
    ("sourcesans", "Source Sans Pro", 20),
    ("avenirnext", "Avenir Next LT Pro", 20),
    ("pmingliu", "PMingLiU", 20),
    ("batang", "Batang", 20),
    ("udnkr", "UD デジタル 教科書体 NK-R", 20),
    ("bizudgothic", "BIZ UDゴシック", 20),
]
SENT = ("The registrar must determine the percentage of care that a person has "
        "for a child during a care period and notify each person concerned. ")


def docx():
    return os.path.join(OUT, "linepitch.docx")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, arm in enumerate(ARMS):
        name, font, sz = arm[0], arm[1], arm[2]
        # 4th field = an explicit spacing element (the `line=0 atLeast` rule the
        # UD worksheet's body paragraphs use, which R55 reads as "natural height,
        # no grid snap")
        spacing = arm[3] if len(arm) > 3 else 'w:before="0" w:after="0" w:line="240" w:lineRule="auto"'
        rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
               '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (font, font, font, sz, sz))
        # ARM MARKER (7pt so the per-arm size filter never picks it up -- at 10pt
        # it joined the 10pt arms' own line set and stretched their spans).
        # An arm whose paragraph overflows onto a second page used to
        # shift every later arm's page by one, and the report read the wrong
        # font's pitch under the right font's name (caught 2026-08-15 when
        # "Lucida Calligraphy" came back with Arial Rounded MT Bold's 1.15727).
        # Each page now carries its own index so the readers can map by marker
        # instead of by position.
        body.append(
            '<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"'
            ' w:hAnsi="Arial"/><w:sz w:val="14"/></w:rPr><w:t>A%02dZ</w:t>'
            "</w:r></w:p>" % ("<w:pageBreakBefore/>" if ai else "", ai))
        body.append(
            '<w:p><w:pPr><w:spacing %s/></w:pPr><w:r>%s'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (spacing, rpr, SENT * (18 if sz <= 24 else 8)))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11907" w:h="16839"/>'
           '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman"'
              ' w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault>"
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>'
              "</w:styles>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms")


def report(per, who):
    print("== %s ==" % who)
    print("%-11s %-18s %6s %7s %6s %9s %9s"
          % ("arm", "font", "size", "lines", "span", "pitch", "em"))
    for ai, arm in enumerate(ARMS):
        name, font, sz = arm[0], arm[1], arm[2]
        ys = per.get(ai) or []
        if len(ys) < 5:
            print("%-11s MISSING (%d lines)" % (name, len(ys)))
            continue
        span = ys[-1] - ys[0]
        pitch = span / (len(ys) - 1)
        print("%-11s %-18s %6.1f %7d %6.2f %9.4f %9.5f"
              % (name, font, sz / 2.0, len(ys), span, pitch, pitch / (sz / 2.0)))


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
    marker_page = {}
    for pi in range(doc.page_count):
        for m in re.finditer(r"A(\d\d)Z", doc[pi].get_text()):
            marker_page.setdefault(int(m.group(1)), pi)
    for ai in range(len(ARMS)):
        pi = marker_page.get(ai)
        if pi is None:
            continue
        # Read BASELINES of the arm's own spans, not line bboxes: the paragraph
        # mark inherits Normal's 10pt Times New Roman, and on the last line that
        # taller span lifts the bbox top by (0.891*10 - 0.891*9) = 0.89pt --
        # spread over 19 gaps that is exactly the -0.046pt/line "deviation" the
        # first run of this probe reported for tnr9.
        want = ARMS[ai][2] / 2.0
        ys = set()
        for bl in doc[pi].get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                for sp in ln["spans"]:
                    if abs(sp["size"] - want) < 0.06 and sp["text"].strip():
                        ys.add(round(sp["origin"][1], 3))
                        break
        per[ai] = sorted(ys)
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "linepitch_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "lp"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    per = {}
    marker_page = {}
    for pi, pg in enumerate(pages):
        for e in pg["elements"]:
            m = re.fullmatch(r"A(\d\d)Z", (e.get("text") or "").strip())
            if m:
                marker_page.setdefault(int(m.group(1)), pi)
    for ai in range(len(ARMS)):
        pi = marker_page.get(ai)
        if pi is None:
            continue
        want = ARMS[ai][2] / 2.0
        per[ai] = sorted({round(e["y"], 3) for e in pages[pi]["elements"]
                          if e.get("type") == "text" and (e.get("text") or "").strip()
                          and abs((e.get("font_size") or 0) - want) < 0.06})
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
