# -*- coding: utf-8 -*-
"""What is the NATURAL line of a `line=0 atLeast` paragraph set in a face whose
hhea box differs from its win box (游ゴシック / 游明朝: hhea 1.602 em with a
0.5 em lineGap, win 1.287 em)?

S1145 derived the CJK natural as `hhea x 83/64` on MS Mincho, where hhea == win,
so the two readings could not be told apart. educational__0214ac95 (three 14pt
empties in "AR P丸ゴシック体E", which Word substitutes with 游ゴシック) drifted
+5.5pt per empty: Oxi 29.09 (hhea) against Word 23.25 (win). Each arm is a text
paragraph followed by an empty one at the same size and face, so consecutive
text baselines in Word's PDF export are 2 x natural (+ the ascent difference).

    python _pb_yunatural_gen.py gen
    python _pb_yunatural_gen.py pdf      # Word truth (COM -> PDF baselines)
    python _pb_yunatural_gen.py oxi      # Oxi, same arms (dump-layout)
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_yunatural")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

# (face, half-point size): a text paragraph + an empty paragraph per arm
ARMS = [
    ("游ゴシック", 21), ("游ゴシック", 24), ("游ゴシック", 28), ("游ゴシック", 32),
    ("游ゴシック Medium", 28), ("游明朝", 28), ("Yu Gothic", 28),
    ("BIZ UDPゴシック", 28), ("メイリオ", 28),
]
DOCX = os.path.join(OUT, "yunatural.docx")


def para(face, sz, text):
    rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>'
           '<w:sz w:val="%d"/></w:rPr>' % (face, face, face, sz))
    run = "<w:r>%s<w:t>%s</w:t></w:r>" % (rpr, text) if text else ""
    return ('<w:p><w:pPr><w:spacing w:line="0" w:lineRule="atLeast"/>%s</w:pPr>%s</w:p>'
            % (rpr, run))


def gen():
    os.makedirs(OUT, exist_ok=True)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
              '<w:kern w:val="2"/><w:sz w:val="21"/></w:rPr></w:rPrDefault>'
              "<w:pPrDefault/></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
              '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
              "</w:styles>")
    body = ""
    for i, (face, sz) in enumerate(ARMS):
        body += para(face, sz, "国国国国%d" % i) + para(face, sz, "")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
           + "><w:body>" + body
           + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
             '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
             '<w:docGrid w:type="lines" w:linePitch="360"/>'
             "</w:sectPr></w:body></w:document>")
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                   'relationships/styles" Target="styles.xml"/></Relationships>')
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), DOCX))


def report(rows, title):
    rows.sort()
    print("== %s: baseline y, delta to previous (= 2 x natural + ascent difference) ==" % title)
    prev = None
    for y, label, face in rows:
        d = ("%.3f" % (y - prev)) if prev is not None else "  -   "
        print("  %8.3f d=%s %-22s %s" % (y, d, face, label))
        prev = y


def pdf():
    import fitz
    import win32com.client as w
    out = DOCX[:-5] + ".word.pdf"
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        d = app.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
        try:
            d.SaveAs2(out, 17)
        finally:
            d.Close(False)
    finally:
        app.Quit()
    rows = []
    for pg in fitz.open(out):
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                chars = [c for sp in l["spans"] for c in sp["chars"] if c["c"].strip()]
                if chars:
                    rows.append((round(chars[0]["origin"][1], 3),
                                 "".join(c["c"] for c in chars)[:8], l["spans"][0]["font"]))
    report(rows, "WORD (PDF)")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    dump = os.path.join(tempfile.gettempdir(), "yunatural.json")
    subprocess.run([GDI, DOCX, os.path.join(tempfile.gettempdir(), "yunat"),
                    "--dump-layout=" + dump], check=True, capture_output=True, env=env)
    seen = set()
    rows = []
    for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
        for e in pg["elements"]:
            if e["type"] == "text" and e.get("text", "").startswith("国"):
                y = round(e["y"] + e.get("text_y_off", 0.0), 3)
                if y not in seen:
                    seen.add(y)
                    rows.append((y, e["text"][:8], "fs=%s" % e.get("font_size")))
    report(rows, "OXI %s" % (envs or "(default)"))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
