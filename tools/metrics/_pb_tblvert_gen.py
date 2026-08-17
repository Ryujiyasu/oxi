# -*- coding: utf-8 -*-
"""Row geometry down the page: where does Word put a table row's boundaries?

The horizontal side of the table box is now pinned (_pb_tblanchor: position,
absorption amount and width all match Word). The vertical side is not, and
tokyoshugyo p24 shows it costing pages: Word's first row there closes at 135.98
with BOTH cells on that line, while Oxi puts the two cells' bottoms at 127.65 and
146.15. Sweep the three things that set a row's height and read Word's own rules.

  cellMar top/bottom -- 0 by default, so it is easy to leave untested
  uneven cells      -- one cell 1 line, its neighbour 2 or 3
  trHeight          -- absent / atLeast / exact, above and below the natural need

    python _pb_tblvert_gen.py gen
    python _pb_tblvert_gen.py pdf      # Word truth
    python _pb_tblvert_gen.py oxi      # Oxi, same arms
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_tblvert")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = "ＭＳ 明朝"
SZ_HP = 21                 # 10.5pt
TOP_TW = 1985
TOP_PT = TOP_TW / 20.0
COMPAT = os.environ.get("OXI_PB_COMPAT", "11")

# (label, cellMar_tb_tw, [lines per cell], trHeight rule or None, trHeight tw)
ARMS = [
    ("even1_cm0", 0, [1, 1], None, 0),
    ("even1_cm108", 108, [1, 1], None, 0),
    ("even1_cm200", 200, [1, 1], None, 0),
    ("uneven12_cm0", 0, [1, 2], None, 0),
    ("uneven12_cm108", 108, [1, 2], None, 0),
    ("uneven13_cm108", 108, [1, 3], None, 0),
    ("uneven31_cm108", 108, [3, 1], None, 0),
    ("atleast_small", 108, [1, 1], "atLeast", 200),
    ("atleast_big", 108, [1, 1], "atLeast", 1200),
    ("exact_small", 108, [2, 2], "exact", 300),
    ("exact_big", 108, [1, 1], "exact", 1200),
    # exact is the one rule Word does not take literally: it adds the cell's TOP
    # margin to the declared height. These three pin what the addend tracks.
    ("exact_cm0", 0, [1, 1], "exact", 1200),
    ("exact_cm200", 200, [1, 1], "exact", 1200),
    ("exact_cm400", 400, [1, 1], "exact", 1200),
    # asymmetric: which of the two margins is the addend?
    ("exact_t400b0", (400, 0), [1, 1], "exact", 1200),
    ("exact_t0b400", (0, 400), [1, 1], "exact", 1200),
]


def docx():
    return os.path.join(OUT, "tblvert.docx")


def para(text):
    return ('<w:p><w:pPr><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/></w:rPr></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (FACE, FACE, FACE, SZ_HP, FACE, FACE, FACE, SZ_HP, SZ_HP, text))


def table(ai, cm, lines, rule, h):
    pr = ["<w:tblPr>", '<w:tblW w:w="0" w:type="auto"/>',
          '<w:tblBorders>' + "".join(
              '<w:%s w:val="single" w:sz="4" w:space="0" w:color="000000"/>' % s
              for s in ("top", "left", "bottom", "right", "insideH", "insideV")) +
          "</w:tblBorders>",
          '<w:tblCellMar><w:top w:w="%d" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
          '<w:bottom w:w="%d" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
          "</w:tblCellMar>" % (cm if isinstance(cm, int) else cm[0],
                               cm if isinstance(cm, int) else cm[1]),
          "</w:tblPr>"]
    trpr = ('<w:trPr><w:trHeight w:val="%d" w:hRule="%s"/></w:trPr>' % (h, rule)) if rule else ""
    cells = []
    for ci, n in enumerate(lines):
        body = "".join(para("R%02dC%d-%d" % (ai, ci, k)) for k in range(n))
        cells.append('<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
                     + body + "</w:tc>")
    return ("<w:tbl>" + "".join(pr) +
            '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
            "<w:tr>" + trpr + "".join(cells) + "</w:tr>"
            # a second, plain row so the first row's BOTTOM rule is unambiguous
            "<w:tr>" + '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
            + para("R%02dTAIL" % ai) + "</w:tc>"
            '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
            + para(".") + "</w:tc>" + "</w:tr>"
            "</w:tbl>")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (label, cm, lines, rule, h) in enumerate(ARMS):
        body.append(para("A%02dZ" % ai))
        body.append(table(ai, cm, lines, rule, h))
        body.append('<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
           '<w:pgMar w:top="%d" w:right="1701" w:bottom="1701" w:left="1701" '
           'w:header="851" w:footer="992" w:gutter="0"/>'
           '<w:docGrid w:type="lines" w:linePitch="360"/>'
           "</w:sectPr></w:body></w:document>" % TOP_TW)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="%d"/></w:rPr></w:style>'
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
    print("%-16s %-5s %-8s %-9s %-9s %-9s %-9s %s"
          % ("arm", "cmTB", "lines", "row_top", "row_bot", "height", "text0_y", "rules"))
    for ai, (label, cm, lines, rule, h) in enumerate(ARMS):
        g = per.get(ai) or {}
        rs = g.get("rules") or []
        t0 = g.get("t0")
        top = rs[0] if len(rs) > 0 else None
        bot = rs[1] if len(rs) > 1 else None
        print("%-16s %-5s %-8s %-9s %-9s %-9s %-9s %s"
              % (label, cm if isinstance(cm, int) else "%d/%d" % cm, "/".join(map(str, lines)),
                 "%.2f" % top if top is not None else "-",
                 "%.2f" % bot if bot is not None else "-",
                 "%.2f" % (bot - top) if (top is not None and bot is not None) else "-",
                 "%.2f" % t0 if t0 is not None else "-",
                 " ".join("%.1f" % r for r in rs[:4])))


def _rules_word(pg):
    ys = []
    for d in pg.get_drawings():
        for it in d["items"]:
            if it[0] == "l" and abs(it[1].y - it[2].y) < 0.3 and abs(it[2].x - it[1].x) > 20:
                ys.append(round((it[1].y + it[2].y) / 2, 2))
            elif it[0] == "re" and it[1].height < 2.0 and it[1].width > 20:
                ys.append(round((it[1].y0 + it[1].y1) / 2, 2))
    out = []
    for y in sorted(ys):
        if not out or y - out[-1] > 0.6:
            out.append(y)
    return out


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
        pg = doc[pi]
        txt = pg.get_text()
        m = re.search(r"R(\d\d)C0-0", txt)
        if not m:
            continue
        ai = int(m.group(1))
        t0 = None
        for b in pg.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                if "R%02dC0-0" % ai in "".join(s["text"] for s in ln["spans"]):
                    t0 = ln["bbox"][1]
        per[ai] = {"rules": _rules_word(pg), "t0": t0}
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "tblvert_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "tv"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    per = {}
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        txt = "".join(e.get("text") or "" for e in pg["elements"] if e["type"] == "text")
        m = re.search(r"R(\d\d)C0-0", txt)
        if not m:
            continue
        ai = int(m.group(1))
        ys = sorted({round(e["y"] + (e.get("h") or 0) / 2.0, 2)
                     for e in pg["elements"]
                     if e["type"] == "border" and (e.get("h") or 0) <= 2.0
                     and (e.get("w") or 0) > 20})
        out_ys = []
        for y in ys:
            if not out_ys or y - out_ys[-1] > 0.6:
                out_ys.append(y)
        rows = {}
        for e in pg["elements"]:
            if e["type"] == "text":
                rows.setdefault(round(e["y"], 1), []).append((e["x"], e.get("text") or ""))
        t0 = None
        for y, v in sorted(rows.items()):
            if "R%02dC0-0" % ai in "".join(t for _, t in sorted(v)):
                t0 = y
                break
        per[ai] = {"rules": out_ys, "t0": t0}
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
