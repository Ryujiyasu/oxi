# -*- coding: utf-8 -*-
"""Where does Word hyphenate an English word, and when does it bother?

educational__00158a7d (0.6348, the worst remaining EN Phase-1 doc after
reports__001f1397) carries `<w:autoHyphenation/>`. Word breaks `Ac-cording`,
`coun-tries`, `poorer coun-` down its pages; Oxi has no hyphenation, spends 1-2
extra lines per page and pushes every page's last paragraph over. 6/619 corpus
docs have the setting.

Two questions, one document:
  (1) the BREAK SET — for a given word, which positions may carry the hyphen
  (2) the TRIGGER — Word only hyphenates when the line would otherwise end
      further left than the hyphenation zone (default 0.25" = 360tw); a word
      that nearly fits is left alone

Each arm is one page holding one paragraph: a filler sentence, then the TARGET
word, under a swept `w:ind w:right` so the line end lands at a different place
inside the target. Reading line 1's tail from the Word PDF gives the break
actually taken at that width; sweeping recovers the whole break set, earliest to
latest, plus the width at which Word stops hyphenating at all.

  python _pb_hyphen_gen.py gen
  python _pb_hyphen_gen.py pdf      # Word truth
  python _pb_hyphen_gen.py oxi      # Oxi's own breaks, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_hyphen")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

PGW, MARG = 12240, 1440          # Letter, 1in margins — the specimen's setup
AVAIL = PGW - 2 * MARG           # 9360tw = 468pt

# The specimen's own vocabulary plus the classic shapes: prefix/suffix splits,
# doubled consonants, -tion, a compound, and a word with no legal break.
TARGETS = [
    "according", "countries", "democracy", "authoritarian", "predisposed",
    "information", "hyphenation", "beautiful", "resource", "through",
]
# Right indent sweep (twips): the line end walks leftward through the target.
INDENTS = [0, 120, 240, 360, 480, 600, 720, 840, 960, 1080, 1200, 1320]
# ★The paragraph is the TARGET repeated: then EVERY line end is a break
# decision about that one word, and the indent sweep only shifts the phase.
# (The first cut used a fixed prose filler with the target last — the target
# always landed mid-line-2 with room to spare and was never the candidate, so
# the only hyphens measured were the filler's.)
REPEAT = 16


def docx():
    return os.path.join(OUT, "hyphen.docx")


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>')


def para(text, indent_r, pbb=False):
    return ('<w:p><w:pPr>%s'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:ind w:right="%d"/><w:jc w:val="left"/>%s</w:pPr>'
            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % ("<w:pageBreakBefore/>" if pbb else "", indent_r, rpr(), rpr(), text))


def arms():
    return [(w, ind) for w in TARGETS for ind in INDENTS]


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (word, ind) in enumerate(arms()):
        body.append(para("M%03d" % ai, 0, pbb=ai > 0))
        # The target is the LAST word, so the break Word takes is unambiguous.
        body.append(para(" ".join([word] * REPEAT) + ".", ind))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="%d" w:h="15840"/>'
           '<w:pgMar w:top="1440" w:right="%d" w:bottom="1440" w:left="%d" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
           % (PGW, MARG, MARG))
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                "<w:autoHyphenation/>"          # exactly as the specimen writes it
                "<w:compat>"
                '<w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
                "</w:compat></w:settings>")
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
              '<w:sz w:val="24"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType='
                    '"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                    "</Types>")
    drels = DRELS.replace("</Relationships>",
                          '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                          'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                          "</Relationships>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(arms()), "arms |", len(TARGETS), "words x",
          len(INDENTS), "indents | avail", AVAIL, "tw")


def report(per, who):
    """per[ai] = [(tail_text, right_pt), ...] for every line of the arm."""
    print("== %s ==" % who)
    print("%-14s %5s  %s" % ("word", "ind", "per-line tail (right, gap)"))
    for ai, (word, ind) in enumerate(arms()):
        lines = per.get(ai)
        if not lines:
            print("%-14s %5d MISSING" % (word, ind))
            continue
        avail = (PGW - MARG - ind) / 20.0
        parts = []
        for tail, right in lines:
            cut = tail.rsplit(" ", 1)[-1] if tail else ""
            parts.append("%s%s(%.1f)" % (cut[-14:], "*" if cut.endswith("-") else "", avail - right))
        print("%-14s %5d  %s" % (word, ind, "  ".join(parts)))


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
    for ai, (_w, ind) in enumerate(arms()):
        if ai >= doc.page_count:
            break
        lines = []
        for bl in doc[ai].get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                t = "".join(s["text"] for s in ln["spans"]).rstrip()
                if t.strip() and not t.strip().startswith("M"):
                    lines.append((round(ln["bbox"][1], 2), t, round(ln["bbox"][2], 2)))
        lines.sort()
        if lines:
            per[ai] = [(t, r) for _y, t, r in lines]
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "hyphen_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "hy"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    per = {}
    for ai, (_w, ind) in enumerate(arms()):
        if ai >= len(pages):
            break
        rows = {}
        for e in pages[ai]["elements"]:
            if e.get("type") != "text":
                continue
            t = e.get("text") or ""
            if not t.strip() or t.strip().startswith("M"):
                continue
            rows.setdefault(round(e["y"], 1), []).append((e["x"], t, e.get("w") or 0))
        if rows:
            per[ai] = []
            for y in sorted(rows):
                cells = sorted(rows[y])
                per[ai].append(("".join(c[1] for c in cells),
                                max(c[0] + c[2] for c in cells)))
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
