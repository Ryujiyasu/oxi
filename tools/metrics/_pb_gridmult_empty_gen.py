# -*- coding: utf-8 -*-
"""How tall is an EMPTY paragraph whose mark is larger than the grid pitch, under
`line=276 auto` on a `docGrid type=lines linePitch=360`?

policies__0353d0b2 p7: three empties with a 14pt mark (line=276 auto, 18pt
grid) between the form label and the title. Word advances ~23.4pt each, Oxi
~38.3 (2 rows + leading) -- +45pt by the title, two paragraphs off the page.
Sweep mark size x line multiple, with the empty paragraphs' mark font both
ways (docDefaults Century/MS Mincho), plus a text-paragraph arm, and read every
paragraph's Information(6) through COM.

    python _pb_gridmult_empty_gen.py gen
    python _pb_gridmult_empty_gen.py com      # Word truth (COM Information(6) per paragraph)
    python _pb_gridmult_empty_gen.py oxi      # Oxi dump per paragraph
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_gridmult_empty")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

# (label, mark sz half-points, w:line, empties?, ascii font override or None)
ARMS = [
    ("s28_l276_empty", 28, 276, True, None),      # the document's case
    ("s28_l240_empty", 28, 240, True, None),
    ("s28_l360_empty", 28, 360, True, None),
    ("s21_l276_empty", 21, 276, True, None),
    ("s24_l276_empty", 24, 276, True, None),
    ("s32_l276_empty", 32, 276, True, None),
    ("s36_l276_empty", 36, 276, True, None),
    ("s28_l276_text", 28, 276, False, None),      # same, with text in the paragraphs
    ("s28_l276_empty_msmincho", 28, 276, True, "ＭＳ 明朝"),   # ascii = MS Mincho too
    ("s28_l276_empty_nogrid", 28, 276, True, "NOGRID"),       # control: no docGrid
]


def docx(label):
    return os.path.join(OUT, "gridmult_%s.docx" % label)


def gen():
    os.makedirs(OUT, exist_ok=True)
    for label, sz, line, empty, asc in ARMS:
        ascii_font = asc if asc and asc != "NOGRID" else "Century"
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
                  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="%s" w:eastAsia="ＭＳ 明朝" w:hAnsi="%s" w:cs="Times New Roman"/>'
                  '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
                  '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>'
                  % (ascii_font, ascii_font))
        mark = '<w:rPr><w:sz w:val="%d"/></w:rPr>' % sz
        body = "<w:p><w:r><w:t>前の行</w:t></w:r></w:p>"
        for i in range(3):
            run = "" if empty else ('<w:r><w:rPr><w:sz w:val="%d"/></w:rPr><w:t>本文%d</w:t></w:r>' % (sz, i))
            body += ('<w:p><w:pPr><w:spacing w:line="%d" w:lineRule="auto"/><w:jc w:val="center"/>%s</w:pPr>%s</w:p>'
                     % (line, mark, run))
        body += "<w:p><w:r><w:t>後の行</w:t></w:r></w:p><w:p><w:r><w:t>末尾</w:t></w:r></w:p>"
        grid = "" if asc == "NOGRID" else '<w:docGrid w:type="lines" w:linePitch="360"/>'
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="737" w:right="849" w:bottom="567" w:left="1134" w:header="851" w:footer="992"/>'
               + grid + "</w:sectPr></w:body></w:document>")
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


def com():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    print("== WORD (COM Information(6), collapsed start): y of each paragraph; d = step ==")
    try:
        for label, sz, line, empty, asc in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                ys = []
                for i in range(1, d.Paragraphs.Count + 1):
                    r = d.Paragraphs(i).Range
                    ys.append(round(d.Range(r.Start, r.Start).Information(6), 2))
                steps = ["%.2f" % (ys[i + 1] - ys[i]) for i in range(len(ys) - 1)]
                print("%-26s sz=%-3s line=%-4s empty=%-5s -> y=%s steps=%s" % (label, sz, line, empty, ys, steps))
            finally:
                d.Close(False)
    finally:
        app.Quit()


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    print("== OXI %s: line tops (text elements) and the 前→後 span ==" % (envs or "(default)"))
    for label, sz, line, empty, asc in ARMS:
        dump = os.path.join(tempfile.gettempdir(), "gridmult_%s.json" % label)
        subprocess.run([GDI, docx(label), os.path.join(tempfile.gettempdir(), "gme"),
                        "--dump-layout=" + dump], check=True, capture_output=True, env=env)
        rows = {}
        for pg in json.load(open(dump, encoding="utf-8"))["pages"]:
            for e in pg["elements"]:
                if e["type"] == "text" and e.get("text", "").strip():
                    rows.setdefault(round(e["y"], 2), []).append(e["text"])
        ys = sorted(rows)
        first = next((y for y in ys if "前" in "".join(rows[y])), None)
        last = next((y for y in ys if "後" in "".join(rows[y])), None)
        print("%-26s tops=%s  前→後=%s" % (label, [(y, "".join(rows[y])[:4]) for y in ys], round(last - first, 2) if first is not None and last is not None else None))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "com":
        com()
    elif cmd == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        gen()
