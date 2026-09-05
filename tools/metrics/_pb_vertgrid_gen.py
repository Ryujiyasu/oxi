# -*- coding: utf-8 -*-
"""What is the COLUMN pitch of vertical (tbRl) text on a `lines` docGrid?

educational__0828836 (landscape, tbRl, linePitch=360, メイリオ 10.5pt body with
9pt translation lines): Word's PDF shows the 9pt columns 18pt apart and the
10.5pt columns ~27pt apart (28/26 alternating); Oxi gives the 10.5pt columns
36 (two cells, the horizontal Meiryo rule) and runs one column over the page.
Sweep face x size in vertical AND horizontal sections and read the line pitch
from Word's PDF export (column x for vertical, baseline y for horizontal).

    python _pb_vertgrid_gen.py gen
    python _pb_vertgrid_gen.py pdf      # Word truth (COM -> PDF)
    python _pb_vertgrid_gen.py oxi      # Oxi dump, same arms
"""
import collections
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_vertgrid")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

FACES = ["メイリオ", "ＭＳ 明朝", "游ゴシック", "ＭＳ ゴシック"]
SIZES = [18, 21, 24, 28]          # half-points: 9 / 10.5 / 12 / 14
ARMS = [("%s_%s_%s" % ("v" if vert else "h", ["meiryo", "msmincho", "yugothic", "msgothic"][fi], sz), vert, face, sz)
        for vert in (True, False) for fi, face in enumerate(FACES) for sz in SIZES]
TEXT = "中納言参り給ひて御扇奉らせ給ふに隆家こそいみじき骨は得て侍れ"


def docx(label):
    return os.path.join(OUT, "vertgrid_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>'
              '<w:kern w:val="2"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
              "<w:pPrDefault/></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
              '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
    for label, vert, face, sz in ARMS:
        rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr>'
               % (face, face, face, sz))
        body = "".join("<w:p><w:pPr>%s</w:pPr><w:r>%s<w:t>%s%d</w:t></w:r></w:p>" % (rpr, rpr, TEXT[:12], i) for i in range(6))
        sect = ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
                '<w:pgMar w:top="1701" w:right="1985" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/>'
                + ('<w:textDirection w:val="tbRl"/>' if vert else "")
                + '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + "><w:body>" + body + sect + "</w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                       'relationships/styles" Target="styles.xml"/></Relationships>')
            z.writestr("word/styles.xml", styles)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def pitches_pdf(path, vert):
    import fitz
    pg = fitz.open(path)[0]
    key = collections.defaultdict(int)
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for sp in l["spans"]:
                for c in sp["chars"]:
                    if c["c"].strip():
                        key[round(c["origin"][0 if vert else 1], 2)] += 1
    ks = sorted(k for k, n in key.items() if n >= 3)
    if vert:
        ks = ks[::-1]
    return [round(abs(ks[i + 1] - ks[i]), 2) for i in range(len(ks) - 1)]


def pdf():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for label, *_ in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                d.SaveAs2(docx(label)[:-5] + ".word.pdf", 17)
            finally:
                d.Close(False)
    finally:
        app.Quit()
    print("== WORD (PDF): line/column pitches between the 6 paragraphs ==")
    for label, vert, face, sz in ARMS:
        print("%-18s %-8s %4.1fpt -> %s" % (label, face, sz / 2, pitches_pdf(docx(label)[:-5] + ".word.pdf", vert)))


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    print("== OXI %s: line/column pitches (dump x for vertical, y for horizontal) ==" % (envs or "(default)"))
    for label, vert, face, sz in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "vertgrid_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "vg"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        key = collections.defaultdict(int)
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and e.get("text", "").strip():
                    key[round(e["x"] if vert else e["y"], 2)] += 1
        ks = sorted(key)
        if vert:
            ks = ks[::-1]
        print("%-18s %-8s %4.1fpt -> %s" % (label, face, sz / 2, [round(abs(ks[i + 1] - ks[i]), 2) for i in range(len(ks) - 1)]))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
