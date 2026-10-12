# -*- coding: utf-8 -*-
"""Which face does Word use for a run whose rFonts names a PostScript name
(ArialMT, TimesNewRomanPSMT, Helvetica...) instead of the family?

blind-G en/correspondence__00a8fa6b: the signature block is «ArialMT» 12pt
(pasted from a PDF). Word sets it on Arial's 13.8pt pitch; Oxi's fallback
metrics give 14.65 (+0.85 per line, «Principal» +3.4 by the third line).

One doc: for each name, three one-word lines «Qq1 / Qq2 / Qq3» at 12pt plus
a width word «Hamburgefonstiv»; read the PDF span font and the baseline pitch.

    python tools/metrics/_pb_psname_gen.py gen
    python tools/metrics/_pb_psname_gen.py word
    python tools/metrics/_pb_psname_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_psname"
NAMES = ["Arial", "ArialMT", "Arial-BoldMT", "Arial-ItalicMT", "ArialNarrow", "Helvetica", "Helvetica-Bold",
         "TimesNewRomanPSMT", "TimesNewRomanPS-BoldMT", "TimesNewRoman", "CourierNewPSMT", "Verdana-Bold",
         "Calibri-Bold", "SymbolMT", "Wingdings-Regular", "NoSuchFontXyz"]
TOK = "Qq"


def para(font, text):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/><w:sz w:val="24"/></w:rPr>'
    return f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{rpr}</w:pPr><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>'


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    body = ""
    for i, n in enumerate(NAMES):
        body += "".join(para(n, f"{TOK}{i:02d}{k} Hamburgefonstiv") for k in range(3))
        body += para("Calibri", f"sep{i:02d}")
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1134" w:bottom="720"'
           ' w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    styles = T.STYLES.replace('w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:eastAsia="MS Gothic" w:cs="MS Gothic"',
                              'w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"')
    with zipfile.ZipFile(OUT / "psname.docx", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", T.CT); z.writestr("_rels/.rels", T.RELS)
        z.writestr("word/_rels/document.xml.rels", T.DRELS); z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("ok", len(NAMES))


def _pitch(lines, i):
    ys = sorted(l["y"] for l in lines if l["t"].startswith(f"{TOK}{i:02d}"))
    return round((ys[-1] - ys[0]) / 2, 2) if len(ys) == 3 else None


def _width(lines, i):
    ls = [l for l in lines if l["t"].startswith(f"{TOK}{i:02d}0")]
    return round(ls[0]["x1"] - ls[0]["x0"], 2) if ls else None


def word():
    import win32com.client, fitz, line_census as L
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        t = os.path.join(tempfile.mkdtemp(), "psname.docx"); shutil.copy(OUT / "psname.docx", t)
        d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
        d.SaveAs2(str((OUT / "psname.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    fonts = {}
    for page in fitz.open(OUT / "psname.pdf"):
        for b in page.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                tx = "".join(s["text"] for s in l["spans"])
                if tx.startswith(TOK):
                    fonts.setdefault(int(tx[2:4]), l["spans"][0]["font"])
    lines = L.word_lines(OUT / "psname.pdf")
    res = {}
    for i, n in enumerate(NAMES):
        res[n] = {"font": fonts.get(i), "pitch": _pitch(lines, i), "width": _width(lines, i)}
        print(f"{n:24s} word face {res[n]['font']!s:28s} pitch {res[n]['pitch']}  width {res[n]['width']}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    lines = L.oxi_lines(os.path.abspath(exe), OUT / "psname.docx")
    for i, n in enumerate(NAMES):
        p, wd = _pitch(lines, i), _width(lines, i)
        flag = "" if p is not None and abs(p - (word[n]["pitch"] or 0)) < 0.3 else "  <-- pitch differs"
        print(f"{n:24s} oxi pitch {p} width {wd}   word {word[n]['pitch']} {word[n]['width']} ({word[n]['font']}){flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
