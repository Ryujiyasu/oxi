# -*- coding: utf-8 -*-
"""What sets the COLUMN advance in vertical (tbRl) writing?

albalunaSS_a6 is the corpus's second-worst document (0.7649) and the whole
residual is one column of slip: its heading carries
`<w:spacing w:line="480" w:lineRule="auto"/>` (double), and Word advances 33.0pt
past it while Oxi advances 26.0pt (= 2 x the 12.8pt docGrid pitch). Body columns
agree exactly at 12.8, so the disagreement is specifically what the multiplier
multiplies.

Horizontal writing's law is "line height = max natural height x multiplier"
([[word_marker_line_height_law]]). If vertical is the same law turned on its
side, the advance is the paragraph's NATURAL column width x multiplier, not the
grid pitch x multiplier -- 33.0 / 2 = 16.5, which is a natural width, not 12.8.
Sweep multiplier x font size x rule and read Word's own column x positions.

    python _pb_vertpitch_gen.py gen
    python _pb_vertpitch_gen.py pdf      # Word truth
    python _pb_vertpitch_gen.py oxi      # Oxi, same arms
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
OUT = os.path.join(REPO, "pipeline_data", "_pb_vertpitch")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = "ＭＳ 明朝"
PITCH = 256                # docGrid linePitch, twips = 12.8pt (as in albalunaSS_a6)
COMPAT = os.environ.get("OXI_PB_COMPAT", "15")

# (label, size_half_points, line value, line rule)
# ★The first run used sz21 (10.5pt) throughout and learned nothing about the
# multiplier: a 10.5pt natural column (14.16) EXCEEDS the 12.8pt grid pitch, so
# every paragraph snapped to 2 cells (measured body pitch 25.68) and the
# multiplier had nowhere to show. Vertical columns grid-snap exactly like
# horizontal lines. Everything below is 8pt (natural < 12.8 = one cell) so the
# marker's advance is the multiplier's own doing; BODY stays 8pt too, so the
# body pitch is the 1-cell reference.
ARMS = [
    ("base_single", 16, None, None),
    ("x1_5", 16, 360, "auto"),
    ("x2", 16, 480, "auto"),
    ("x3", 16, 720, "auto"),
    ("exact200", 16, 200, "exact"),
    ("exact400", 16, 400, "exact"),
    ("exact600", 16, 600, "exact"),
    ("atleast200", 16, 200, "atLeast"),
    ("atleast400", 16, 400, "atLeast"),
    # ★Composition arms (2026-08-21, the 047ff775/01535587 pcd=-6/-5 hunt):
    # sz21 = 10.5pt MS明朝, natural column 14.16 > pitch 12.8 → the run-1
    # observation says single snaps to 2 cells (25.68). These discriminate
    # model A (advance = mult × cells × pitch) from model B (advance =
    # ceil(natural×mult/pitch) × pitch):
    #   x1.5: A=38.4  B=25.6      x2: A=51.2  B=38.4     x3: A=76.8 B=51.2
    ("cell2_single", 21, None, None),
    ("cell2_x1_5", 21, 360, "auto"),
    ("cell2_x2", 21, 480, "auto"),
    ("cell2_x3", 21, 720, "auto"),
    ("cell2_exact200", 21, 200, "exact"),
    ("cell2_exact400", 21, 400, "exact"),
    ("cell2_atl200", 21, 200, "atLeast"),
    ("cell2_atl400", 21, 400, "atLeast"),
    # ceil tolerance boundary: 9.5pt natural = 12.81 vs pitch 12.80 (horizontal
    # S752 uses ceil((nat-0.5)/pitch) — does vertical share the 0.5pt window?);
    # 10pt natural = 13.49 (clears any sub-0.7 tolerance → 2 cells).
    ("boundary_sz19", 19, None, None),
    ("boundary_sz20", 20, None, None),
]
COLS = 3                   # columns of body text after the swept paragraph


def docx():
    return os.path.join(OUT, "vertpitch.docx")


def para(text, sz, ppr=""):
    return ('<w:p><w:pPr>%s<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/></w:rPr></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (ppr, FACE, FACE, FACE, sz, FACE, FACE, FACE, sz, sz, text))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (label, sz, line, rule) in enumerate(ARMS):
        body.append(para("A%02dZ" % ai, 16,
                         "<w:pageBreakBefore/>" if ai else ""))
        ppr = ('<w:spacing w:line="%d" w:lineRule="%s"/>' % (line, rule)) if line else ""
        body.append(para("M%02dマーカー" % ai, sz, ppr))
        for k in range(COLS):
            body.append(para("B%02d%d本文の列です" % (ai, k), 16))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="8392" w:h="11907" w:code="11"/>'
           '<w:pgMar w:top="1134" w:right="737" w:bottom="794" w:left="1191" '
           'w:header="227" w:footer="170" w:gutter="0"/>'
           '<w:textDirection w:val="tbRl"/>'
           '<w:docGrid w:type="lines" w:linePitch="%d"/>'
           "</w:sectPr></w:body></w:document>" % PITCH)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>'
              "</w:styles>" % (FACE, FACE, FACE))
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
    print("wrote", docx(), len(ARMS), "arms; grid pitch %.2fpt; compat %s"
          % (PITCH / 20.0, COMPAT))


def report(per, who):
    print("== %s == (grid pitch %.2fpt)" % (who, PITCH / 20.0))
    print("%-16s %-6s %-11s %-9s %-9s %-9s %s"
          % ("arm", "sz_pt", "line", "marker_x", "body0_x", "advance", "body pitch"))
    for ai, (label, sz, line, rule) in enumerate(ARMS):
        g = per.get(ai) or {}
        mx, bx = g.get("m"), g.get("b")
        bp = g.get("bp")
        ax = g.get("a")
        print("%-16s %-6.1f %-11s %-8s %-8s %-8s %-8s %-8s %s"
              % (label, sz / 2.0,
                 ("%d %s" % (line, rule)) if line else "single",
                 "%.2f" % ax if ax is not None else "-",
                 "%.2f" % mx if mx is not None else "-",
                 "%.2f" % bx if bx is not None else "-",
                 "%.2f" % (ax - mx) if (ax is not None and mx is not None) else "-",
                 "%.2f" % (mx - bx) if (mx is not None and bx is not None) else "-",
                 " ".join("%.2f" % v for v in (bp or []))))


def _collect(spans_per_page):
    per = {}
    for spans in spans_per_page:
        ai = None
        # ★Vertical flows RIGHT to LEFT, and a marker that needs more than one
        # column emits several spans. Take the RIGHTMOST (= the first column);
        # assigning span-by-span kept the LAST one and made the exact/atLeast
        # arms unreadable.
        for x, t in spans:
            m = re.search(r"M(\d\d)", t)
            if m:
                ai = int(m.group(1))
                d = per.setdefault(ai, {})
                d["m"] = max(d.get("m", -1e9), x)
                d["mn"] = d.get("mn", 0) + 1
        if ai is None:
            continue
        for x, t in spans:
            if re.search(r"A%02dZ" % ai, t):
                d = per.setdefault(ai, {})
                d["a"] = max(d.get("a", -1e9), x)
        bs = sorted({round(x, 2) for x, t in spans if re.search(r"B%02d\d" % ai, t)},
                    reverse=True)
        if bs:
            per[ai]["b"] = bs[0]
            per[ai]["bp"] = [round(bs[i] - bs[i + 1], 2) for i in range(len(bs) - 1)]
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
    pages = []
    for pi in range(doc.page_count):
        spans = []
        for b in doc[pi].get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                for s in ln["spans"]:
                    if s["text"].strip():
                        spans.append((s["bbox"][0], s["text"]))
        pages.append(spans)
    report(_collect(pages), "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "vertpitch_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "vp"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = []
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        pages.append([(e["x"], e.get("text") or "")
                      for e in pg["elements"] if e["type"] == "text"])
    report(_collect(pages), "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
