# -*- coding: utf-8 -*-
"""Where does Word hang a paragraph's TOP border when the line has extra leading?

technical__002c1ffa's footer is an empty bordered paragraph, a 2x3 table and an
empty paragraph, bottom-anchored 170.1pt above the page edge.  Both engines put
the table at the same y (Word 627.35 / Oxi 627.20), so their footer stacks agree
-- but Word draws the separator rule at 616.39 and Oxi at 612.45.

3.94pt is close to the extra leading that document's Normal style forces on
every line: `w:line=260 w:lineRule=atLeast` = 13pt against 8pt Times New Roman's
natural 9.199, i.e. 3.80pt of leading.  Oxi hangs the border off the LINE BOX
top; if Word hangs it off the TEXT top instead, the leading sits above the
border and the 3.94 is explained.

The arms separate the three candidate anchors -- box top, text top, and
first-baseline -- by varying how much leading there is and where it comes from
(atLeast, exact, a line multiplier), with an auto-spaced arm as the zero.

  python _pb_bdratleast_gen.py gen
  python _pb_bdratleast_gen.py pdf      # Word truth
  python _pb_bdratleast_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_bdratleast")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

# (arm, line twips, rule, run half-points, has text, spacing before twips)
ARMS = [
    ("auto8", 240, "auto", 16, True, 0),
    ("atleast260_8", 260, "atLeast", 16, True, 0),
    ("atleast400_8", 400, "atLeast", 16, True, 0),
    ("atleast260_empty", 260, "atLeast", 16, False, 0),      # the footer's own shape
    ("exact260_8", 260, "exact", 16, True, 0),
    ("auto8_before120", 240, "auto", 16, True, 120),
    ("atleast260_11", 260, "atLeast", 22, True, 0),          # leading only 0.35pt
    ("double8", 480, "auto", 16, True, 0),                   # multiplier leading
]
TXT = "Compilation No. 51"


def docx():
    return os.path.join(OUT, "bdratleast.docx")


def marker(tag, ai, brk):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"'
            ' w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr><w:t>%s%02d</w:t>'
            "</w:r></w:p>" % ("<w:pageBreakBefore/>" if brk else "", tag, ai))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (name, line, rule, sz, has_text, before) in enumerate(ARMS):
        body.append(marker("M", ai, ai > 0))
        run = ('<w:r><w:rPr><w:rFonts w:ascii="Times New Roman"'
               ' w:hAnsi="Times New Roman"/><w:sz w:val="%d"/></w:rPr>'
               "<w:t>%s</w:t></w:r>" % (sz, TXT)) if has_text else ""
        body.append(
            '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="6" w:space="1"'
            ' w:color="auto"/></w:pBdr><w:spacing w:before="%d" w:after="0"'
            ' w:line="%d" w:lineRule="%s"/><w:rPr><w:rFonts'
            ' w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="%d"/></w:rPr></w:pPr>%s</w:p>'
            % (before, line, rule, sz, run))
        body.append(marker("E", ai, False))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11907" w:h="16839"/>'
           '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman"'
              ' w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
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
    print("%-18s %8s %8s %8s %9s %9s"
          % ("arm", "M_top", "rule_y", "E_top", "rule-M", "E-rule"))
    for ai, arm in enumerate(ARMS):
        g = per.get(ai)
        if not g or g.get("rule") is None:
            print("%-18s %s" % (arm[0], "NO RULE" if g else "MISSING"))
            continue
        m, r, e = g.get("m"), g["rule"], g.get("e")
        print("%-18s %8.2f %8.2f %8.2f %9.2f %9.2f"
              % (arm[0], m or 0, r, e or 0, (r - m) if m else 0, (e - r) if e else 0))


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
    for ai in range(min(len(ARMS), doc.page_count)):
        g = {"rule": None}
        for dr in doc[ai].get_drawings():
            if dr["rect"].width > 100:
                g["rule"] = round(dr["rect"].y0, 2)
        for bl in doc[ai].get_text("dict")["blocks"]:
            for ln in bl.get("lines", []):
                for sp in ln["spans"]:
                    t = sp["text"].strip()
                    # markers are Arial 10; report their TOP so both engines'
                    # numbers mean the same thing
                    top = round(sp["origin"][1] - sp["ascender"] * sp["size"], 2)
                    if t.startswith("M"):
                        g["m"] = top
                    elif t.startswith("E"):
                        g["e"] = top
                    elif t:
                        g["text_top"] = top
                        g["base"] = round(sp["origin"][1], 2)
        per[ai] = g
    report(per, "WORD")
    for ai, arm in enumerate(ARMS):
        g = per.get(ai) or {}
        if g.get("base"):
            print("   %-18s text_top=%.2f baseline=%.2f  rule=%.2f  base-rule=%.2f"
                  % (arm[0], g["text_top"], g["base"], g["rule"], g["base"] - g["rule"]))


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "bdratleast_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "ba"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    per = {}
    for ai in range(min(len(ARMS), len(pages))):
        g = {"rule": None}
        for e in pages[ai]["elements"]:
            if e["type"] != "text":
                if (e.get("w") or 0) > 100:
                    g["rule"] = round(e["y"], 2)
                continue
            t = (e.get("text") or "").strip()
            if t.startswith("M"):
                g["m"] = round(e["y"], 2)
            elif t.startswith("E"):
                g["e"] = round(e["y"], 2)
            elif t:
                g.setdefault("text_top", round(e["y"], 2))
        per[ai] = g
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
