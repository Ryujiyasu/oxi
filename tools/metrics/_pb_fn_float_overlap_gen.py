# -*- coding: utf-8 -*-
"""Does a floating shape that reaches past the body keep footnotes off the page?

en/forms__0068e33b: page-sized frame rectangles (wrapNone, behind text) and a
text box's clip record reach below the bottom margin; Oxi took them as body
content, found no room for the footnote area and dropped the note (Word sets
it at the page bottom). Faithful slices of ja/educational__20aa108d: the
paragraph that references footnote 1 gets a 100 x 400pt rectangle anchored to
it (column x 300, paragraph y 0), so it runs ~250pt past the bottom margin:
    none      no rectangle
    wrapnone  wrapNone, behind text
    square    wrapSquare
    topbot    wrapTopAndBottom
Read: page 1's footnote rule y (the 144pt stroke) and whether the note's text
(«本» below the rule) is on page 1.

    python tools/metrics/_pb_fn_float_overlap_gen.py   (PB_OXI=<exe>[;<exe>] adds Oxi)
"""
import json, os, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/educational/20aa108db65e9ee3.docx"
OUT = REPO / "pipeline_data" / "_pb_fn_float_overlap"
ARMS = ["none", "wrapnone", "square", "topbot"]
WRAP = {"wrapnone": "<wp:wrapNone/>", "square": '<wp:wrapSquare wrapText="bothSides"/>', "topbot": "<wp:wrapTopAndBottom/>"}


def rect(wrap, behind):
    emu = 12700
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
            f'relativeHeight="251659264" behindDoc="{1 if behind else 0}" locked="0" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>{300 * emu}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{100 * emu}" cy="{400 * emu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap}'
            f'<wp:docPr id="901" name="Probe Rect"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            f'<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{100 * emu}" cy="{400 * emu}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>'
            f'<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>'
            f'</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def variant(name):
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    doc = parts["word/document.xml"].decode("utf-8")
    if name != "none":
        i = doc.find("<w:footnoteReference")
        k = doc.rfind("<w:r>", 0, i)
        k2 = doc.rfind("<w:r ", 0, i)
        k = max(k, k2)
        doc = doc[:k] + rect(WRAP[name], name == "wrapnone") + doc[k:]
    parts["word/document.xml"] = doc.encode("utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return p


def read_word(pdf):
    import pymupdf
    doc = pymupdf.open(pdf)
    pg = doc[0]
    rule = [round(d["rect"].y0, 2) for d in pg.get_drawings() if abs(d["rect"].width - 144) < 1 and d["rect"].height < 1.6]
    rule = rule[0] if rule else None
    note = any(c["c"] == "本" and rule and c["origin"][1] > rule for b in pg.get_text("rawdict")["blocks"]
               for l in b.get("lines", []) for s in l["spans"] for c in s["chars"])
    return f"pages {len(doc)} rule {rule} note-on-p1 {note}"


def read_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json"); lay = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    subprocess.run([exe, str(docx), os.path.join(t, "q"), "--dump-layout=" + lay], capture_output=True)
    pages = json.load(open(lay, encoding="utf-8"))["pages"]
    rule = [round(e["y"], 2) for e in pages[0]["elements"] if abs(e.get("w", 0) - 144) < 0.6 and e.get("h", 9) < 1.6]
    rule = rule[0] if rule else None
    note = any(c["char"] == "本" and rule and c["baseline"] > rule for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"])
    return f"pages {len(pages)} rule {rule} note-on-p1 {note}"


if __name__ == "__main__":
    exes = [e for e in os.environ.get("PB_OXI", "").split(";") if e]
    for name in os.environ.get("PB_ARMS", ",".join(ARMS)).split(","):
        p = variant(name); pdf = p.with_suffix(".pdf")
        if not pdf.exists():
            import win32com.client
            w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
            try:
                d = w.Documents.Open(str(p), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
            finally:
                w.Quit()
        print(f"{name:9} W {read_word(pdf)}" + "".join(f"   O {read_oxi(e, p)}" for e in exes), flush=True)
