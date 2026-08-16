# -*- coding: utf-8 -*-
"""With widowControl OFF, what does Word still refuse to leave at a page break?

Two observations that cannot both be explained by a line count:

  * b837 (S282) declares `<w:widowControl w:val="0"/>` in Normal, yet Word's
    rendering protects a 7-line paragraph -- Oxi, following the XML, split it and
    ran a page ahead from that point on.
  * d77a pi=46 (S283) is a 4-line paragraph in the same situation and Word does
    NOT protect it; forcing protection there regressed the document.
  * reports__11b5f886 (2026-08-16) is a 2-line paragraph, widowControl=0
    inherited from Normal, and Word moves BOTH lines to the next page.

S283 reconciled the first two with a ">= 5 lines" threshold, which the third
falsifies.  The alternative that fits all three is a split-shape rule rather than
a size rule: Word refuses to leave a paragraph's FIRST line alone at the bottom
of a page (an orphan) but will happily send its LAST line alone to the top of the
next one (a widow).  b837's 7-line and our 2-line cases both break as orphans;
d77a's 4-line case breaks as a widow.

Each arm sets the cursor with filler lines so that exactly K lines of the test
paragraph fit before the page bottom, then reads which page each line landed on.

  python _pb_orphan_gen.py gen
  python _pb_orphan_gen.py pdf      # Word truth
  python _pb_orphan_gen.py oxi      # Oxi, same arms
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_orphan")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

# Body: 11pt Times New Roman, single spacing -> 12.649pt per line.
# Page 842 - top 70.9 - bottom 70.9 = 700.2 of body, i.e. 55 lines per page.
FILL_LINE = "The registrar must determine the percentage of care for a child. "
# (arm, total lines in the test paragraph, lines that fit before the break)
ARMS = [
    ("p2_fit1", 2, 1),   # orphan: first line alone at the bottom
    ("p3_fit1", 3, 1),   # orphan
    ("p3_fit2", 3, 2),   # widow: last line alone at the top
    ("p4_fit1", 4, 1),   # orphan
    ("p4_fit3", 4, 3),   # widow
    ("p5_fit2", 5, 2),   # neither: 2 + 3
    ("p7_fit1", 7, 1),   # orphan, the b837 shape
]
LINES_PER_PAGE = 55


def docx():
    return os.path.join(OUT, "orphan.docx")


def para(text, sz_hp=22, extra=""):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman"'
            ' w:hAnsi="Times New Roman"/><w:sz w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (extra, sz_hp, text))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (name, total, fit) in enumerate(ARMS):
        # marker + fillers so the test paragraph starts exactly `fit` lines
        # above the page bottom
        body.append(para("A%02dZ" % ai, 14,
                         "<w:pageBreakBefore/>" if ai else ""))
        for k in range(LINES_PER_PAGE - fit - 1):
            body.append(para("filler %02d %02d" % (ai, k)))
        # the test paragraph: `total` lines, each tagged so its page is readable
        txt = " ".join("T%02dL%02d %s" % (ai, i, FILL_LINE) for i in range(total))
        body.append(para(txt))
        body.append(para("E%02dZ" % ai, 14))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11907" w:h="16839"/>'
           '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    # widowControl OFF in the default style -- the whole point of the probe
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman"'
              ' w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr>'
              '<w:rPr><w:sz w:val="22"/></w:rPr></w:style>'
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
    print("%-9s %6s %5s %-26s %s" % ("arm", "lines", "fit", "line -> page", "verdict"))
    for ai, (name, total, fit) in enumerate(ARMS):
        g = per.get(ai) or {}
        pages = [g.get(i) for i in range(total)]
        if any(p is None for p in pages):
            print("%-9s %6d %5d MISSING %s" % (name, total, fit, pages))
            continue
        first = pages[0]
        split_at = next((i for i, p in enumerate(pages) if p != first), None)
        if split_at is None:
            verdict = "MOVED WHOLE" if fit < total else "kept whole"
        else:
            verdict = "split %d+%d" % (split_at, total - split_at)
        print("%-9s %6d %5d %-26s %s" % (name, total, fit, str(pages), verdict))


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
        for m in re.finditer(r"T(\d\d)L(\d\d)", doc[pi].get_text()):
            per.setdefault(int(m.group(1)), {}).setdefault(int(m.group(2)), pi + 1)
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "orphan_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "or"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    per = {}
    for pi, pg in enumerate(pages):
        txt = "".join(e.get("text") or "" for e in pg["elements"] if e["type"] == "text")
        for m in re.finditer(r"T(\d\d)L(\d\d)", txt):
            per.setdefault(int(m.group(1)), {}).setdefault(int(m.group(2)), pi + 1)
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
