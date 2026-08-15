# -*- coding: utf-8 -*-
"""What line height does Word use for a font that is not installed?

creative__00d0925f sets its whole body in `Avenir Book` 11pt — a Mac font absent
from this machine (its fontTable also names Avenir Next, Didot and Baskerville).
Word fits the document on ONE page at a 12.0pt line pitch; Oxi falls back to its
default face and lays it out at 13.0, needs a second page, and the doc's 1.0000
score turns out to be an artifact of which paragraphs happened to land on page 1.

The question is which metrics Word uses once it substitutes: the substitute
face's own, or the ones declared for the missing font in `word/fontTable.xml`
(`w:panose1`, `w:sig`, `w:family`, `w:pitch`). Each arm is one page holding a
two-line paragraph in one font, so the pitch is the y difference between the two
lines. Arms cover the specimen's font with and without its fontTable entry,
a second missing font, and installed controls.

  python _pb_fontsub_gen.py gen
  python _pb_fontsub_gen.py pdf      # Word truth
  python _pb_fontsub_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_fontsub")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

SPECIMEN = os.path.join(REPO, "pipeline_data", "docx_corpus", "en", "creative",
                        "00d0925fd44848ef.docx")

# (name, font, size half-points, fontTable entry?)
ARMS = [
    ("a_avenir_ft", "Avenir Book", 22, True),     # the specimen's exact setup
    ("b_avenir_noft", "Avenir Book", 22, False),  # same font, no fontTable entry
    ("c_didot_ft", "Didot", 22, True),
    ("d_bogus_noft", "Zzqx Nonesuch", 22, False),
    ("e_calibri", "Calibri", 22, True),           # installed control
    ("f_arial", "Arial", 22, True),               # installed control
    ("g_times", "Times New Roman", 22, True),     # installed control
    ("h_avenir_16", "Avenir Book", 32, True),     # size sweep: 16pt
    # ★what does Word fall back to when it knows NOTHING about the font? Arm d
    # (no fontTable entry) came out Cambria, but that is one shape. Oxi's
    # unknown-font fallback is Calibri and several shipped rules were calibrated
    # on it, so the fallback is only worth changing if Word's choice holds
    # across the declared family/pitch — these arms declare a fontTable entry
    # for the same bogus name with each family value.
    ("i_bogus_swiss", "Zzqx Swiss", 22, "swiss"),
    ("j_bogus_roman", "Zzqx Roman", 22, "roman"),
    ("k_bogus_modern", "Zzqx Modern", 22, "modern"),
    ("l_bogus_auto", "Zzqx Auto", 22, "auto"),
]
# A line long enough to wrap exactly once at the probe's column width.
LINE = ("The quick brown fox jumps over the lazy dog while the quick brown fox "
        "jumps over the lazy dog and then keeps running past the lazy dog again.")


def docx():
    return os.path.join(OUT, "fontsub.docx")


def para(font, sz, text, pbb=False):
    rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
           '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (font, font, font, sz, sz))
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/>%s</w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t>'
            "</w:r></w:p>" % ("<w:pageBreakBefore/>" if pbb else "", rpr, rpr, text))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (name, font, sz, _ft) in enumerate(ARMS):
        body.append(para("Arial", 20, "M%02d" % ai, pbb=ai > 0))
        body.append(para(font, sz, LINE))
        body.append(para("Arial", 20, "E%02d" % ai))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/>'
              "</w:rPr></w:rPrDefault>"
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    # the specimen's own fontTable, so the declared panose/sig for Avenir Book
    # and Didot are byte-identical to the real document's
    ftbl = zipfile.ZipFile(SPECIMEN).read("word/fontTable.xml").decode("utf-8")
    extra = "".join(
        '<w:font w:name="%s"><w:panose1 w:val="00000000000000000000"/>'
        '<w:charset w:val="00"/><w:family w:val="%s"/><w:pitch w:val="variable"/>'
        "</w:font>" % (font, fam)
        for _n, font, _sz, fam in ARMS if isinstance(fam, str))
    ftbl = ftbl.replace("</w:fonts>", extra + "</w:fonts>").encode("utf-8")
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/fontTable.xml" ContentType='
                    '"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
                    "</Types>")
    drels = DRELS.replace("</Relationships>",
                          '<Relationship Id="rIdFT" Type="http://schemas.openxmlformats.org/'
                          'officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'
                          "</Relationships>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/fontTable.xml", ftbl)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms")


def report(per, who):
    print("== %s ==" % who)
    print("%-14s %-18s %-5s %8s %10s" % ("arm", "font", "size", "pitch", "lines"))
    for ai, (name, font, sz, _ft) in enumerate(ARMS):
        got = per.get(ai)
        if not got:
            print("%-14s %-18s %-5.1f MISSING" % (name, font, sz / 2.0))
            continue
        ys = got
        pitch = (ys[-1] - ys[0]) / max(1, len(ys) - 1) if len(ys) > 1 else 0.0
        print("%-14s %-18s %-5.1f %8.2f %10d" % (name, font, sz / 2.0, pitch, len(ys)))


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
    for ai in range(len(ARMS)):
        if ai >= doc.page_count:
            break
        ys, fonts = [], set()
        for bl in doc[ai].get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t.startswith(("M", "E")) or not t:
                    continue
                ys.append(round(ln["bbox"][1], 2))
                for s in ln["spans"]:
                    fonts.add(s["font"])
        ys.sort()
        if ys:
            per[ai] = ys
            print("   %-14s Word drew it in: %s" % (ARMS[ai][0], sorted(fonts)))
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "fontsub_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "fs"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    per = {}
    for ai in range(len(ARMS)):
        if ai >= len(pages):
            break
        ys = set()
        for e in pages[ai]["elements"]:
            if e.get("type") != "text":
                continue
            t = (e.get("text") or "").strip()
            if not t or t.startswith(("M", "E")):
                continue
            ys.add(round(e["y"], 2))
        if ys:
            per[ai] = sorted(ys)
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
