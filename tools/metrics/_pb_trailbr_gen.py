# -*- coding: utf-8 -*-
"""Does a TRAILING <w:br/> open a line Word then leaves empty?

Found while attributing educational__00161422's p12/p13 boundary. Its
paragraph 290 ends with a bare `<w:br/>` as its LAST run, and the Word PDF puts
27.6pt between that paragraph's last line (y=476.70) and the next paragraph's
first (y=504.30) where the line pitch is 14.6 -- i.e. the trailing break costs a
whole empty line. Oxi runs a uniform pitch there, fits one line more per page
and lands the row boundary a line late. `<w:br/>` already parses to '\\n'
(ooxml.rs), so the loss is downstream, and S927 fixed exactly this class for
`<w:cr>` ("dropping it collapsed a genuine trailing empty line").

Arms, each a paragraph P followed by a probe paragraph Q; the readout is the
GAP from P's last line to Q's first line:

  plain      P has no break                      -> one pitch
  trail1     P ends with one <w:br/>             -> two pitches if Word opens it
  trail2     P ends with two <w:br/>             -> three
  mid        P has a <w:br/> in the middle       -> control, one pitch
  trailsp    P ends with <w:br/> then a space    -> is a space enough to hold it

    python _pb_trailbr_gen.py gen
    python _pb_trailbr_gen.py pdf   # Word truth
    python _pb_trailbr_gen.py oxi   # Oxi
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_trailbr")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

PGW, PGH, MARG = 12240, 15840, 1440
ARMS = ["plain", "trail1", "trail2", "mid", "trailsp", "cell", "cellplain"]


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="24"/></w:rPr>')


def run(text="", br=0, tail_space=False):
    out = ""
    if text:
        out += "<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>" % (rpr(), text)
    for _ in range(br):
        out += "<w:r>%s<w:br/></w:r>" % rpr()
    if tail_space:
        out += "<w:r>%s<w:t xml:space=\"preserve\"> </w:t></w:r>" % rpr()
    return out


def para_raw(inner, pbb=False):
    return ('<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/></w:pPr>%s</w:p>'
            % ("<w:pageBreakBefore/>" if pbb else "", inner))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, arm in enumerate(ARMS):
        body.append(para_raw(run("M%02d" % ai), pbb=ai > 0))
        if arm == "plain":
            inner = run("a%dP one" % ai)
        elif arm == "trail1":
            inner = run("a%dP one" % ai, br=1)
        elif arm == "trail2":
            inner = run("a%dP one" % ai, br=2)
        elif arm == "mid":
            inner = run("a%dP one" % ai, br=1) + run("a%dP two" % ai)
        else:
            inner = run("a%dP one" % ai, br=1, tail_space=True)
        if arm in ("cell", "cellplain"):
            # same pair, but inside a one-cell table -- 00161422's context
            inner = run("a%dP one" % ai, br=0 if arm == "cellplain" else 1)
            body.append(
                '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
                '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/>'
                '<w:left w:val="single" w:sz="4" w:color="000000"/>'
                '<w:bottom w:val="single" w:sz="4" w:color="000000"/>'
                '<w:right w:val="single" w:sz="4" w:color="000000"/></w:tblBorders>'
                "</w:tblPr>"
                '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>'
                '<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>'
                + para_raw(inner) + para_raw(run("a%dQ" % ai))
                + "</w:tc></w:tr></w:tbl>")
            continue
        body.append(para_raw(inner))
        body.append(para_raw(run("a%dQ" % ai)))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="%d" w:h="%d"/>'
           '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
           % (PGW, PGH, MARG, MARG, MARG, MARG))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
              '<w:sz w:val="24"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/></w:style></w:styles>')
    with zipfile.ZipFile(os.path.join(OUT, "trailbr.docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", os.path.join(OUT, "trailbr.docx"), len(ARMS), "arms")


def docx():
    return os.path.join(OUT, "trailbr.docx")


def report(pos, who):
    print("== %s ==" % who)
    print("%-9s %-10s %-10s %-8s %s" % ("arm", "P_last_y", "Q_y", "gap", "pitches"))
    for ai, arm in enumerate(ARMS):
        pl = pos.get("a%dP one" % ai)
        if arm == "mid":
            pl = pos.get("a%dP two" % ai) or pl
        q = pos.get("a%dQ" % ai)
        if not pl or not q:
            print("%-9s MISSING (P=%s Q=%s)" % (arm, pl, q))
            continue
        gap = q[1] - pl[1]
        print("%-9s %-10.2f %-10.2f %-8.2f %.2f" % (arm, pl[1], q[1], gap, gap / 13.8))


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
    pos = {}
    for pi in range(doc.page_count):
        for bl in doc[pi].get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if t:
                    pos.setdefault(t, (pi, round(ln["bbox"][1], 2)))
    report(pos, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "trailbr_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "tbr"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pos = {}
    for pi, pg in enumerate(json.load(open(out, encoding="utf-8"))["pages"]):
        # join same-y fragments so "a0P one" is found even when split into runs
        rows = {}
        for e in pg["elements"]:
            if e.get("type") != "text":
                continue
            t = e.get("text") or ""
            if t.strip():
                rows.setdefault(round(e["y"], 2), []).append((e["x"], t))
        for y, frags in rows.items():
            s = "".join(t for _x, t in sorted(frags)).strip()
            if s:
                pos.setdefault(s, (pi, y))
    report(pos, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
