# -*- coding: utf-8 -*-
"""Which font size is the character grid's DEFAULT when docDefaults has no w:sz?

reference__0ea3ec86 (Normal sz=22, docDefaults without w:sz, linesAndChars
charSpace=2048): Word advances its 11pt body at 11.0pt (21 chars in a 235.6pt
column); Oxi takes the Normal size as the grid default, pitch 11 + 0.5 = 11.5,
and holds 20. Two readings agree with 11.0: the grid default is the Japanese
built-in 10.5pt (pitch 11.0, the 11pt run expanded to the pitch) or 10pt
(pitch 10.5, the run at its natural 11). Sweep the Normal size with the
docDefaults size absent, and once with it present, and read the advance of
「国」 in Word's own PDF export.

    python _pb_griddefault_gen.py gen
    python _pb_griddefault_gen.py pdf      # Word truth (COM -> PDF)
    python _pb_griddefault_gen.py oxi      # Oxi, same arms
"""
import collections
import json
import os
import statistics
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_griddefault")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

MINCHO = "ＭＳ 明朝"
# (label, docDefaults sz half-points or None, Normal sz half-points or None, run sz or None, charSpace)
ARMS = [
    ("dd_none_n22", None, 22, None, 2048),
    ("dd_none_n21", None, 21, None, 2048),
    ("dd_none_n20", None, 20, None, 2048),
    ("dd_none_n24", None, 24, None, 2048),
    ("dd_none_none", None, None, None, 2048),     # no size anywhere: the built-in default
    ("dd_none_n22_r21", None, 22, 21, 2048),      # Normal 11, run 10.5
    ("dd22_n22", 22, 22, None, 2048),
    ("dd21_n22", 21, 22, None, 2048),
    ("dd20_n22", 20, 22, None, 2048),
    ("dd_none_n22_cs0", None, 22, None, None),    # control: no charSpace
    ("dd_none_n22_cs3194", None, 22, None, 3194),
]


def docx(label):
    return os.path.join(OUT, "griddefault_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                    "</Types>")
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
             'relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
             'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
             'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             "</Relationships>")
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS + ">"
                '<w:characterSpacingControl w:val="compressPunctuation"/>'
                "<w:compat><w:useFELayout/>"
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
                "</w:compat>"
                '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>')
    for label, dd, nsz, rsz, cs in ARMS:
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  "<w:docDefaults><w:rPrDefault><w:rPr>"
                  '<w:rFonts w:ascii="Century" w:eastAsia="%s" w:hAnsi="Century"/>' % MINCHO
                  + ('<w:sz w:val="%d"/>' % dd if dd else "")
                  + "</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>"
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>'
                  + ('<w:rPr><w:sz w:val="%d"/></w:rPr>' % nsz if nsz else "")
                  + "</w:style></w:styles>")
        rpr = '<w:rPr><w:rFonts w:hint="eastAsia"/>' + ('<w:sz w:val="%d"/>' % rsz if rsz else "") + "</w:rPr>"
        text = "国" * 60 + "、" + "国" * 20
        para = "<w:p><w:r>%s<w:t>%s</w:t></w:r></w:p>" % (rpr, text)
        grid = '<w:docGrid w:type="linesAndChars" w:linePitch="411"%s/>' % (' w:charSpace="%d"' % cs if cs is not None else "")
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + para
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1304" w:right="1021" w:bottom="1134" w:left="1021" w:header="680" w:footer="567"/>'
               + grid + "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels", drels)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def advances_pdf(path):
    import fitz
    pg = fitz.open(path)[0]
    adv = []
    sizes = set()
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            chars = [c for sp in l["spans"] for c in sp["chars"] if c["c"].strip()]
            sizes.update(round(sp["size"], 2) for sp in l["spans"])
            xs = [c["origin"][0] for c in chars]
            full = xs and xs[-1] > 480.0
            if not full:
                continue   # justified full lines only? no: take the unjustified LAST line instead
    # last line (unjustified) advances
    rows = []
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            chars = [c for sp in l["spans"] for c in sp["chars"] if c["c"] == "国"]
            if len(chars) > 3:
                rows.append((chars[0]["origin"][1], [chars[i + 1]["origin"][0] - chars[i]["origin"][0] for i in range(len(chars) - 1)]))
    rows.sort()
    last = rows[-1][1] if rows else []
    first = rows[0][1] if rows else []
    return sorted(sizes), (round(statistics.median(first), 2) if first else None), (round(statistics.median(last), 2) if last else None), len(rows)


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
    print("== WORD (PDF): span sizes, 国 advance on the first (justified) and last (natural) line, lines ==")
    for label, dd, nsz, rsz, cs in ARMS:
        sizes, a1, a2, n = advances_pdf(docx(label)[:-5] + ".word.pdf")
        print("%-20s dd=%-4s normal=%-4s run=%-4s cs=%-5s -> sizes=%s first=%s last=%s lines=%d" % (label, dd, nsz, rsz, cs, sizes, a1, a2, n))


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    print("== OXI %s: 国 width (dump w) on first / last line, lines ==" % (envs or "(default)"))
    for label, dd, nsz, rsz, cs in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "griddefault_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "gdf"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        rows = collections.defaultdict(list)
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and e.get("text") == "国":
                    rows[round(e["y"], 1)].append(e)
        ys = sorted(rows)
        f = statistics.median(e["w"] for e in rows[ys[0]]) if ys else None
        l = statistics.median(e["w"] for e in rows[ys[-1]]) if ys else None
        fs = sorted(set(round(e["font_size"], 2) for y in ys for e in rows[y]))
        print("%-20s -> fs=%s first=%.2f last=%.2f lines=%d" % (label, fs, f, l, len(ys)))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
