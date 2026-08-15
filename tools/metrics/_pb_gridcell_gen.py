# -*- coding: utf-8 -*-
"""How far can a line exceed the docGrid pitch before Word gives it a second cell?

The CJK half of the S1142 font sweep is held on this question.  Adding real UD
Digi Kyokasho metrics turns educational__0214ac95 -- `docGrid type="lines"
linePitch="360"`, body 85% UD -- from Word's 2 pages into 3, and the arithmetic
points at one boundary: a 12pt line at that face's CJK-inflated 1.5em is exactly
18.0pt, the grid pitch itself.  Oxi takes a second cell there; Word does not.

So sweep the line height across the pitch in fine steps and read the pitch Word
actually lays down.  Times New Roman's natural is 1.14990em, so size S gives
1.1499*S: sizes 15..17 walk the height from 17.25 to 19.55 across the 18pt cell.
Each arm is one page holding a paragraph long enough to wrap several times, and
the answer is the y-difference between its lines -- 18 for one cell, 36 for two.

A second block repeats the sweep with a CJK face (MS Mincho, whose 1.0em natural
is inflated by 83/64) because the held entries are CJK and the grid rule may
read the CJK metric rather than the composed line height.

  python _pb_gridcell_gen.py gen
  python _pb_gridcell_gen.py pdf      # Word truth
  python _pb_gridcell_gen.py oxi      # Oxi, same arms
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_gridcell")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

PITCH = 360          # twips = 18pt, the pitch the regressing document uses
# (arm, font, half-points): TNR 1.14990em -> natural = 1.1499 * size
ARMS = [
    ("tnr15", "Times New Roman", 30),      # 17.25
    ("tnr15h", "Times New Roman", 31),     # 17.82
    ("tnr156", "Times New Roman", 31.3),   # 17.99  (w:sz is half-points; 31.3 -> 31)
    ("tnr16", "Times New Roman", 32),      # 18.40
    ("tnr165", "Times New Roman", 33),     # 18.97
    ("tnr17", "Times New Roman", 34),      # 19.55
    # CJK side: MS Mincho natural 1.0em, Word draws it at 83/64 = 1.296875em, so
    # size S gives 1.296875*S -- 13.5pt lands on 17.51, 14pt on 18.16.
    ("min13", "MS Mincho", 26),            # 16.86
    ("min135", "MS Mincho", 27),           # 17.51
    ("min14", "MS Mincho", 28),            # 18.16
    ("min15", "MS Mincho", 30),            # 19.45
    # Tighten the bracket around the pitch itself: Word puts 17.82 in one cell
    # and 18.16 in two, so the boundary is somewhere in between. These three
    # step 17.95 -> 18.17 using faces whose natural x a half-point size lands
    # there (Trebuchet 1.16113, Segoe UI 1.32995, Cambria 1.17248).
    ("segoe135", "Segoe UI", 27),          # 17.954
    ("treb155", "Trebuchet MS", 31),       # 17.997  -- the exact-pitch case
    ("camb155", "Cambria", 31),            # 18.173
    # R55 says `line=0 atLeast` means "natural height, no grid snap" -- but that
    # was derived in a document with NO docGrid. The UD worksheet's four
    # spacing-bearing body paragraphs carry exactly this rule INSIDE a
    # type=lines grid, and they are the only body lines whose height changes
    # when the CJK metrics land. So ask the question in the grid.
    ("min16_at0", "MS Mincho", 32, 'w:before="0" w:after="0" w:line="0" w:lineRule="atLeast"'),
    ("min13_at0", "MS Mincho", 26, 'w:before="0" w:after="0" w:line="0" w:lineRule="atLeast"'),
    ("tnr16_at0", "Times New Roman", 32, 'w:before="0" w:after="0" w:line="0" w:lineRule="atLeast"'),
]
SENT = ("The registrar must determine the percentage of care that a person has "
        "for a child during a care period and notify each person concerned. ")
SENT_JA = "この文書は行グリッドの一マスに収まる高さの上限を測るための本文です。"


def docx():
    return os.path.join(OUT, "gridcell.docx")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, arm in enumerate(ARMS):
        name, font, sz = arm[0], arm[1], int(arm[2])
        spacing = arm[3] if len(arm) > 3 else ('w:before="0" w:after="0"'
                                               ' w:line="240" w:lineRule="auto"')
        cjk = font.startswith("MS ")
        rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
               '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
               % (font, font, font, sz, sz))
        # 7pt Arial marker: identifies the arm's page even if a paragraph spills
        body.append(
            '<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"'
            ' w:hAnsi="Arial"/><w:sz w:val="14"/></w:rPr><w:t>A%02dZ</w:t>'
            "</w:r></w:p>" % ("<w:pageBreakBefore/>" if ai else "", ai))
        body.append(
            '<w:p><w:pPr><w:spacing %s/></w:pPr><w:r>%s'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (spacing, rpr, (SENT_JA * 12) if cjk else (SENT * 6)))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11907" w:h="16839"/>'
           '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
           'w:header="720" w:footer="720" w:gutter="0"/>'
           '<w:docGrid w:type="lines" w:linePitch="%d"/></w:sectPr></w:body></w:document>'
           % PITCH)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:eastAsia="MS Mincho"'
              ' w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="21"/></w:rPr></w:style>'
              "</w:styles>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms, pitch", PITCH / 20.0, "pt")


def report(per, who):
    print("== %s ==" % who)
    print("%-9s %-18s %6s %7s %9s %8s" % ("arm", "font", "size", "lines", "pitch", "cells"))
    for ai, arm in enumerate(ARMS):
        name, font, sz = arm[0], arm[1], arm[2]
        ys = per.get(ai) or []
        if len(ys) < 3:
            print("%-9s MISSING (%d lines)" % (name, len(ys)))
            continue
        pitch = (ys[-1] - ys[0]) / (len(ys) - 1)
        print("%-9s %-18s %6.1f %7d %9.3f %8.2f"
              % (name, font, int(sz) / 2.0, len(ys), pitch, pitch / (PITCH / 20.0)))


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
    page_of = {}
    for pi in range(doc.page_count):
        for m in re.finditer(r"A(\d\d)Z", doc[pi].get_text()):
            page_of.setdefault(int(m.group(1)), pi)
    per = {}
    for ai, arm in enumerate(ARMS):
        sz = arm[2]
        pi = page_of.get(ai)
        if pi is None:
            continue
        want = int(sz) / 2.0
        ys = set()
        for bl in doc[pi].get_text("dict")["blocks"]:
            for ln in bl.get("lines", []):
                for sp in ln["spans"]:
                    if abs(sp["size"] - want) < 0.3 and sp["text"].strip():
                        ys.add(round(sp["origin"][1], 3))
                        break
        per[ai] = sorted(ys)
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "gridcell_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "gc"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    page_of = {}
    for pi, pg in enumerate(pages):
        for e in pg["elements"]:
            m = re.fullmatch(r"A(\d\d)Z", (e.get("text") or "").strip())
            if m:
                page_of.setdefault(int(m.group(1)), pi)
    per = {}
    for ai, arm in enumerate(ARMS):
        sz = arm[2]
        pi = page_of.get(ai)
        if pi is None:
            continue
        want = int(sz) / 2.0
        per[ai] = sorted({round(e["y"], 3) for e in pages[pi]["elements"]
                          if e.get("type") == "text" and (e.get("text") or "").strip()
                          and abs((e.get("font_size") or 0) - want) < 0.3})
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
