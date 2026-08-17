# -*- coding: utf-8 -*-
"""How does Word split a table row across a page boundary?

_pb_tblvert pinned the row heights on a page (13 of 16 arms already matched; the
exact-rule addend was the one gap, S1164). What it does not cover is the case
tokyoshugyo p24-29 actually hits: a row that STARTS near the page bottom and
continues on the next page. There Word closes the first row at 135.98 with both
cells on that line while Oxi puts the two cells' bottoms at 127.65 and 146.15,
and the whole tail of the document inherits -17pt.

Arms sweep the distance from the row's top to the content bottom (a spacer
paragraph moves it), how many lines each of the two cells holds, and cantSplit.
Read from Word: how many of each cell's lines stay on the first page.

    python _pb_rowsplit_gen.py gen
    python _pb_rowsplit_gen.py pdf      # Word truth
    python _pb_rowsplit_gen.py oxi      # Oxi, same arms
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_rowsplit")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = "ＭＳ 明朝"
SZ_HP = 21                 # 10.5pt
TOP_TW = 1985
BOT_TW = 1701
PGH = 16838
COMPAT = os.environ.get("OXI_PB_COMPAT", "11")
FILLERS = int(os.environ.get("OXI_PB_FILL", "33"))  # body lines before the table
LINE_PT = 18.0             # docGrid pitch for the body paragraphs

# (label, spacer_tw, [lines per cell], cantSplit)
LINES_SETS = [[4, 4], [4, 2], [2, 4], [6, 1]]
SPACERS = [0, 60, 120, 180, 240, 300, 360, 420, 480]
ARMS = [("s%d_%s" % (sp, "x".join(map(str, ls))), sp, ls, False)
        for ls in LINES_SETS for sp in SPACERS]
ARMS += [("cantsplit_s%d" % sp, sp, [4, 4], True) for sp in (120, 240, 360)]


def docx():
    return os.path.join(OUT, "rowsplit.docx")


def para(text, ppr=""):
    return ('<w:p><w:pPr>%s<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/></w:rPr></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (ppr, FACE, FACE, FACE, SZ_HP, FACE, FACE, FACE, SZ_HP, SZ_HP, text))


def table(ai, lines, cantsplit):
    pr = ("<w:tblPr>" + '<w:tblW w:w="0" w:type="auto"/>' +
          '<w:tblBorders>' + "".join(
              '<w:%s w:val="single" w:sz="4" w:space="0" w:color="000000"/>' % s
              for s in ("top", "left", "bottom", "right", "insideH", "insideV")) +
          "</w:tblBorders>" +
          '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
          '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
          "</w:tblCellMar>" + "</w:tblPr>")
    trpr = "<w:trPr><w:cantSplit/></w:trPr>" if cantsplit else ""
    cells = []
    for ci, n in enumerate(lines):
        body = "".join(para("R%02dC%dL%d" % (ai, ci, k)) for k in range(n))
        cells.append('<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
                     + body + "</w:tc>")
    return ("<w:tbl>" + pr +
            '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
            "<w:tr>" + trpr + "".join(cells) + "</w:tr></w:tbl>")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (label, sp, lines, cs) in enumerate(ARMS):
        body.append(para("A%02dZ" % ai,
                         "<w:pageBreakBefore/>" if ai else ""))
        for k in range(FILLERS):
            body.append(para("うめ%02d-%02d" % (ai, k)))
        if sp:
            body.append(para("s", '<w:spacing w:before="0" w:after="0"'
                                  ' w:line="%d" w:lineRule="exact"/>' % sp))
        body.append(table(ai, lines, cs))
        body.append(para("E%02dZ" % ai))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="%d" w:code="9"/>'
           '<w:pgMar w:top="%d" w:right="1701" w:bottom="%d" w:left="1701" '
           'w:header="851" w:footer="992" w:gutter="0"/>'
           '<w:docGrid w:type="lines" w:linePitch="360"/>'
           "</w:sectPr></w:body></w:document>" % (PGH, TOP_TW, BOT_TW))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr>'
              '<w:rPr><w:sz w:val="%d"/></w:rPr></w:style>'
              "</w:styles>" % (FACE, FACE, FACE, SZ_HP))
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + NS +
                '><w:compat><w:compatSetting w:name="compatibilityMode"'
                ' w:uri="http://schemas.microsoft.com/office/word"'
                ' w:val="%s"/></w:compat></w:settings>' % COMPAT)
    ct = CT.replace("</Types>",
                    '<Override PartName="/word/settings.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
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
    print("wrote", docx(), len(ARMS), "arms; compat", COMPAT)


def report(per, who):
    print("== %s ==" % who)
    print("%-14s %-7s %-8s %-6s %-11s %-11s %s"
          % ("arm", "spacer", "lines", "split", "c0 on p1/p2", "c1 on p1/p2", "verdict"))
    for ai, (label, sp, lines, cs) in enumerate(ARMS):
        g = per.get(ai)
        if not g:
            print("%-14s %-7.1f %-8s MISSING" % (label, sp / 20.0, "x".join(map(str, lines))))
            continue
        c0, c1 = g.get(0, {}), g.get(1, {})
        p0 = sorted(set(c0.values()))
        n0 = [sum(1 for v in c0.values() if v == p) for p in p0]
        p1 = sorted(set(c1.values()))
        n1 = [sum(1 for v in c1.values() if v == p) for p in p1]
        moved = "whole-move" if (len(p0) == 1 and len(p1) == 1
                                 and min(p0) == min(p1) and min(p0) > g.get("apage", 0)) else ""
        v = moved or ("split" if (len(p0) > 1 or len(p1) > 1) else "fits")
        print("%-14s %-7.1f %-8s %-6s %-11s %-11s %s"
              % (label, sp / 20.0, "x".join(map(str, lines)),
                 "cs" if cs else "-",
                 "/".join(map(str, n0)), "/".join(map(str, n1)), v))


def _collect(pagetexts):
    per = {}
    for pi, t in enumerate(pagetexts):
        for m in re.finditer(r"A(\d\d)Z", t):
            per.setdefault(int(m.group(1)), {})["apage"] = pi + 1
        for m in re.finditer(r"R(\d\d)C(\d)L(\d)", t):
            ai, ci, li = int(m.group(1)), int(m.group(2)), int(m.group(3))
            per.setdefault(ai, {}).setdefault(ci, {})[li] = pi + 1
    return per


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
    report(_collect([doc[i].get_text() for i in range(doc.page_count)]), "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "rowsplit_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "rs"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    texts = ["".join(e.get("text") or "" for e in pg["elements"] if e["type"] == "text")
             for pg in json.load(open(out, encoding="utf-8"))["pages"]]
    report(_collect(texts), "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
