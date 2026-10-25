# -*- coding: utf-8 -*-
"""Where does Word put the text of a footnote whose style says snapToGrid=0, and
its separator rule?

blind-G ja/educational__20aa108d (lines 360 grid, footnote text style = Normal +
snapToGrid 0 + jc left, the note runs 9pt ＭＳ 明朝, one note): Word's note line box
ends at the bottom margin (666.45) as in Oxi, but Word's baseline is 663.84 (the
11.67 box with its text centred) and Oxi's 665.03 (+1.19); the separator rule is
649.80 in Word and 642.16 in Oxi.

Faithful slices of the real package, one property changed per arm:
    asis      as it is
    snap1     the footnote-text style's snapToGrid=0 removed (the note snaps)
    sz21      the note at 10.5pt (sz 21) instead of 9pt
    nogrid    the section's docGrid removed
    sz14      the note at 7pt
    nosup     the footnoteRef run without its superscript character style
    noref     no footnoteRef run (the note text alone)
Read: the note's baseline («本»), its reference «1», the separator rule y
(the 144pt-wide drawing), the page's content bottom.

    python tools/metrics/_pb_fn_snap0_gen.py         (PB_OXI=<exe>[;<exe>] adds Oxi)
"""
import json, os, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/educational/20aa108db65e9ee3.docx"
OUT = REPO / "pipeline_data" / "_pb_fn_snap0"


def variant(name):
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    doc = parts["word/document.xml"].decode("utf-8")
    sty = parts["word/styles.xml"].decode("utf-8")
    fn = parts["word/footnotes.xml"].decode("utf-8")
    if name == "snap1":
        i = sty.find('w:styleId="ad"')
        j = sty.find("</w:style>", i)
        sty = sty[:i] + sty[i:j].replace('<w:snapToGrid w:val="0"/>', "") + sty[j:]
    elif name in ("sz21", "sz14"):
        v = name[2:]
        fn = fn.replace('<w:sz w:val="18"/><w:szCs w:val="18"/>', f'<w:sz w:val="{v}"/><w:szCs w:val="{v}"/>')
    elif name == "nosup":
        # the reference run without its superscript (rStyle ae = vertAlign superscript)
        fn = fn.replace('<w:rStyle w:val="ae"/>', "")
    elif name == "noref":
        # no footnoteRef run at all (the note text alone)
        import re as _re
        fn = _re.sub(r'<w:r><w:rPr><w:rStyle w:val="ae"/>.*?<w:footnoteRef/></w:r>', "", fn, flags=_re.S)
    elif name == "nogrid":
        doc = doc.replace('<w:docGrid w:type="lines" w:linePitch="360"/>', "")
    parts.update({"word/document.xml": doc.encode("utf-8"), "word/styles.xml": sty.encode("utf-8"),
                  "word/footnotes.xml": fn.encode("utf-8")})
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return p


def read_word(pdf):
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    base = ref = None
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"] == "本" and c["origin"][1] > 600 and base is None:
                        base = c["origin"][1]
                    if c["c"] == "1" and c["origin"][1] > 600 and s["size"] < 8 and ref is None:
                        ref = c["origin"][1]
    rule = [round(d["rect"].y0, 2) for d in pg.get_drawings() if abs(d["rect"].width - 144) < 1 and d["rect"].y0 > 550]
    return base, ref, rule[0] if rule else None


def read_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json"); lay = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    subprocess.run([exe, str(docx), os.path.join(t, "q"), "--dump-layout=" + lay], capture_output=True)
    base = ref = None
    for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"]:
        if c["char"] == "本" and c["baseline"] > 600 and base is None:
            base = c["baseline"]
        if c["char"] == "1" and c["baseline"] > 600 and ref is None:
            ref = c["baseline"]
    rule = [round(e["y"], 2) for e in json.load(open(lay, encoding="utf-8"))["pages"][0]["elements"]
            if abs(e.get("w", 0) - 144) < 1 and e["y"] > 550]
    return base, ref, rule[0] if rule else None


def fmt(v):
    base, ref, rule = v
    f = lambda x: "%.2f" % x if x is not None else "-"
    return f"note {f(base)} ref {f(ref)} rule {f(rule)}"


if __name__ == "__main__":
    import win32com.client
    exes = [e for e in os.environ.get("PB_OXI", "").split(";") if e]
    for name in os.environ.get("PB_ARMS", "asis,snap1,sz21,sz14,nogrid").split(","):
        p = variant(name); pdf = p.with_suffix(".pdf")
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(str(p), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
        finally:
            w.Quit()
        print(f"{name:7} W {fmt(read_word(pdf))}" + "".join(f"   O {fmt(read_oxi(e, p))}" for e in exes), flush=True)
