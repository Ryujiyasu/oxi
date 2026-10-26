# -*- coding: utf-8 -*-
"""Footnote lines under `line=240 atLeast` (ja/reports__62ceebc6).

The notes are Times New Roman 9pt, snapToGrid 0, `spacing line=240
lineRule=atLeast` (12pt floor), in a linesAndChars 350 section. Word sets every
note line 12.00 apart; Oxi reserved 12 per line but laid the lines out 10.35
apart (Times 9 natural). Faithful slices: the body is one paragraph that
references footnote 2, the package is otherwise the real one.
    asis       as is
    snap1      the note paragraph's snapToGrid 0 removed
    nogrid     the section's docGrid removed
    exact12    lineRule exact (12pt)
    atleast16  line=320 atLeast
Read: the note's text baselines (Times 9) and their pitch.

    python tools/metrics/_pb_fn_atleast_gen.py   (PB_OXI=<exe>[;<exe>] adds Oxi)
"""
import json, os, re, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/reports/62ceebc6327acb28.docx"
OUT = REPO / "pipeline_data" / "_pb_fn_atleast"
ARMS = ["asis", "snap1", "nogrid", "exact12", "atleast16"]


def variant(name):
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    doc = parts["word/document.xml"].decode("utf-8")
    sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)[-1]
    if name == "nogrid":
        sect = re.sub(r"<w:docGrid [^>]*/>", "", sect)
    i = doc.find("<w:body>") + len("<w:body>")
    j = doc.rfind("</w:body>")
    body = ('<w:p><w:r><w:t>X</w:t></w:r><w:r><w:rPr><w:rStyle w:val="af0"/></w:rPr>'
            '<w:footnoteReference w:id="2"/></w:r></w:p>' + sect)
    parts["word/document.xml"] = (doc[:i] + body + doc[j:]).encode("utf-8")
    fn = parts["word/footnotes.xml"].decode("utf-8")
    a = fn.find('w:id="2"')
    b = fn.find("</w:footnote>", a)
    note = fn[a:b]
    if name == "snap1":
        note = note.replace('<w:snapToGrid w:val="0"/>', "", 1)
    elif name == "exact12":
        note = note.replace('w:line="240" w:lineRule="atLeast"', 'w:line="240" w:lineRule="exact"')
    elif name == "atleast16":
        note = note.replace('w:line="240" w:lineRule="atLeast"', 'w:line="320" w:lineRule="atLeast"')
    parts["word/footnotes.xml"] = (fn[:a] + note + fn[b:]).encode("utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, data in parts.items():
            w.writestr(n, data)
    return p


def pitch(ys):
    ys = sorted(ys)
    return f"{[round(y, 2) for y in ys]} pitch {[round(b - a, 2) for a, b in zip(ys, ys[1:])]}"


def read_word(pdf):
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    return pitch({round(s["origin"][1], 2) for b in pg.get_text("dict")["blocks"] for l in b.get("lines", [])
                  for s in l["spans"] if abs(s["size"] - 9) < 0.1 and s["text"].strip()})


def read_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    pg = json.load(open(g, encoding="utf-8"))["pages"][0]
    return pitch({round(c["baseline"], 2) for c in pg["glyphs"] if abs(c["font_size"] - 9) < 0.1 and c["char"].strip()})


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
        print(f"{name:10} W {read_word(pdf)}", flush=True)
        for e in exes:
            print(f"{'':10} O {read_oxi(e, p)}", flush=True)
