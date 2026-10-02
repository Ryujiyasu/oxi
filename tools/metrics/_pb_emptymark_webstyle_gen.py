# -*- coding: utf-8 -*-
"""Probe (2026-10-02): the height of an EMPTY paragraph between two text lines
when its paragraph mark carries its own ascii/hAnsi font (Open Sans 12 in a
Times New Roman NormalWeb style) and the paragraph says textAlignment=baseline.

Host: en/policies 0098c921 (Droid Serif -> Cambria via altName, Open Sans).
Body per arm: «Line A» (Cambria italic 12, NormalWeb, spacing 0), the empty
paragraph (mark font per arm), «Line B» (Open Sans 11).  Arms:
   mark font {opensans, tnr(style)} x textAlignment {baseline, none}
Read: baseline pitch A->B (Word PDF span origins / Oxi --dump-glyphs).

    python tools/metrics/_pb_emptymark_webstyle_gen.py <renderer.exe> [oxi-only]
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/policies/0098c9210b9c74e3.docx"
OUT = REPO / "tests/fixtures/emptymark_webstyle"
ARMS = [(f, t) for f in ("opensans", "tnr") for t in ("baseline", "none")]
SP = '<w:spacing w:before="0" w:beforeAutospacing="0" w:after="0" w:afterAutospacing="0"/>'


def name(a):
    return f"{a[0]}_{a[1]}"


def body(font, ta):
    tal = '<w:textAlignment w:val="baseline"/>' if ta == "baseline" else ""
    mark = '<w:rPr><w:rFonts w:ascii="Open Sans" w:hAnsi="Open Sans"/></w:rPr>' if font == "opensans" else ""
    return (f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>{SP}{tal}</w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Droid Serif" w:hAnsi="Droid Serif"/><w:i/></w:rPr><w:t>Line A</w:t></w:r></w:p>'
            f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>{SP}{tal}{mark}</w:pPr></w:p>'
            f'<w:p><w:pPr><w:pStyle w:val="NormalWeb"/>{SP}{tal}</w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Open Sans" w:hAnsi="Open Sans"/><w:sz w:val="22"/></w:rPr><w:t>Line B</w:t></w:r></w:p>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    b0 = doc.index("<w:body>") + len("<w:body>")
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    for a in ARMS:
        xml = doc[:b0] + body(*a) + sect + "</w:body></w:document>"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"{name(a)}.docx").write_bytes(buf.getvalue())


def word_read():
    import fitz, win32com.client
    res = {}
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for a in ARMS:
            tmp = os.path.join(tempfile.mkdtemp(), name(a) + ".docx"); shutil.copy(OUT / f"{name(a)}.docx", tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            pdf = tmp[:-5] + ".pdf"; d.SaveAs2(pdf, 17); d.Close(0)
            got = {}
            for b in fitz.open(pdf)[0].get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    for s in l["spans"]:
                        if s["text"].startswith("Line A"): got["A"] = round(s["origin"][1], 2)
                        if s["text"].startswith("Line B"): got["B"] = round(s["origin"][1], 2)
            res[name(a)] = got
    finally:
        w.Quit()
    return res


def oxi_read(exe):
    sys.path.insert(0, str(REPO / "tools/metrics"))
    from line_census import win_ascent
    res = {}
    for a in ARMS:
        t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
        subprocess.run([exe, str(OUT / f"{name(a)}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True)
        got = {}
        gl = json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]
        # «Line A» / «Line B»: the 'A' / 'B' glyphs
        for g in gl:
            if g["char"] in "AB" and g["char"] not in got:
                got[g["char"]] = round(g["top"] + win_ascent(g["font_family"]) * g["font_size"], 2)
        res[name(a)] = got
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    gen()
    O = oxi_read(os.path.abspath(sys.argv[1]))
    W = word_read() if len(sys.argv) < 3 else {name(a): {} for a in ARMS}
    for a in ARMS:
        n = name(a); wv, ov = W[n], O[n]
        f = lambda d: f"A->B {d['B'] - d['A']:6.2f}" if "A" in d and "B" in d else str(d)
        print(f"{n:18} Word {f(wv)}   | Oxi {f(ov)}")
