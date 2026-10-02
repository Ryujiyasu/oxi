# -*- coding: utf-8 -*-
"""Probe (2026-10-02): where does Word measure a floating table's tblpY when
vertAnchor is absent and the default header is tall (pushes the body top
below the top margin)?

Host: en/forms 005e0208 (header2 = 59pt picture, margin.top 72, tblpY 2745 =
137.25pt, no vertAnchor). Arms rewrite the package:
  asis      : the document as it is
  nohdrpic  : the header picture paragraph emptied (body top = margin)
  vmargin   : vertAnchor="margin" written explicitly
  vtext     : vertAnchor="text"
  vpage     : vertAnchor="page"
Read: the first horizontal table rule's y in the Word PDF, and Oxi's first
'border' element y (--dump-layout).

    python tools/metrics/_pb_tblpy_anchor_gen.py <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/forms/005e0208c0e35c16.docx"
OUT = REPO / "tests/fixtures/tblpy_anchor"
ARMS = ["asis", "nohdrpic", "vmargin", "vtext", "vpage"]


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    for a in ARMS:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml" and a.startswith("v"):
                    x = data.decode("utf-8")
                    x = x.replace('<w:tblpPr ', f'<w:tblpPr w:vertAnchor="{a[1:]}" ', 1)
                    data = x.encode("utf-8")
                if item.filename == "word/header2.xml" and a == "nohdrpic":
                    x = data.decode("utf-8")
                    x = re.sub(r"<w:r>(?:(?!</w:r>).)*<w:drawing>.*?</w:drawing>.*?</w:r>", "", x, flags=re.S)
                    data = x.encode("utf-8")
                zout.writestr(copy.copy(item), data)
        (OUT / f"{a}.docx").write_bytes(buf.getvalue())


def word_rules():
    import fitz, win32com.client
    res = {}
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for a in ARMS:
            tmp = os.path.join(tempfile.mkdtemp(), a + ".docx"); shutil.copy(OUT / f"{a}.docx", tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            pdf = tmp[:-5] + ".pdf"; d.SaveAs2(pdf, 17); d.Close(0)
            pg = fitz.open(pdf)[0]
            ys = sorted(set(round(dr["rect"].y0, 2) for dr in pg.get_drawings() if dr["rect"].width > 100 and dr["rect"].height < 3))
            first = None
            for b in pg.get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    for s in l["spans"]:
                        if s["text"].strip() and (first is None or s["origin"][1] < first[0]):
                            first = (round(s["origin"][1], 2), s["text"][:16])
            res[a] = {"rule": ys[0] if ys else None, "first_text": first}
    finally:
        w.Quit()
    return res


def oxi_rules(exe):
    res = {}
    for a in ARMS:
        t = tempfile.mkdtemp(); out = os.path.join(t, "l.json")
        subprocess.run([exe, str(OUT / f"{a}.docx"), os.path.join(t, "p"), "110", "--dump-layout=" + out], capture_output=True)
        els = json.load(open(out, encoding="utf-8"))["pages"][0]["elements"]
        ys = sorted(set(round(e["y"], 2) for e in els if e.get("type") == "border" and e.get("w", 0) > 100 and e.get("h", 9) < 2))
        txt = [(round(e["y"] + e.get("text_y_off", 0), 2), e["text"][:16]) for e in els if e.get("type") == "text" and e.get("text", "").strip()]
        res[a] = {"rule": ys[0] if ys else None, "first_text": min(txt) if txt else None}
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    gen()
    W = word_rules(); O = oxi_rules(os.path.abspath(sys.argv[1]))
    for a in ARMS:
        print(f"{a:10} Word rule {W[a]['rule']}  first {W[a]['first_text']}   | Oxi rule {O[a]['rule']}  first {O[a]['first_text']}")
