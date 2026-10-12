# -*- coding: utf-8 -*-
"""How does Word advance characters of various sizes on a linesAndChars grid
whose charSpace is NEGATIVE?

blind-G ja/reports__58dca3b2 (docGrid linesAndChars linePitch=290 charSpace=-2271,
docDefault 10.5pt, 2 columns): Word advances the 12 / 11 / 10pt runs of the
title line at 11.6 / 10.5 / 9.5 (≈ fs − 0.55 = fs + charSpace/4096) but the
5pt «・・・き り と り・・・» line at 5.04 (= the em, no shrink); Oxi gave the
first line of that paragraph 28.3 per character and wrapped it, overflowing a
column.

One doc per charSpace in {-2271, -1000, 0, 1453}: paragraphs of HG丸ｺﾞｼｯｸM-PRO
at sz 10 / 12 / 16 / 20 / 21 / 24 / 28 (5..14pt), each «あいうえおかきくけこ»;
read the per-character advance from the PDF.

    python tools/metrics/_pb_negcharspace_gen.py gen
    python tools/metrics/_pb_negcharspace_gen.py word
    python tools/metrics/_pb_negcharspace_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_negcharspace"
CS = {"m2271": -2271, "m1000": -1000, "z0": 0, "p1453": 1453}
SZS = (10, 12, 16, 20, 21, 24, 28)
TEXT = "あいうえおかきくけこ"
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
          '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')


def para(sz):
    rpr = f'<w:rPr><w:rFonts w:ascii="HG丸ｺﾞｼｯｸM-PRO" w:eastAsia="HG丸ｺﾞｼｯｸM-PRO" w:hAnsi="HG丸ｺﾞｼｯｸM-PRO" w:hint="eastAsia"/><w:sz w:val="{sz}"/></w:rPr>'
    return f'<w:p><w:pPr>{rpr}</w:pPr><w:r>{rpr}<w:t>{TEXT}</w:t></w:r></w:p>'


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="840"/>'
                '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat>'
                '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    for arm, cs in CS.items():
        body = "".join(para(sz) for sz in SZS)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="454" w:right="284" w:bottom="284" w:left="284"'
               f' w:header="851" w:footer="992" w:gutter="0"/><w:docGrid w:type="linesAndChars" w:linePitch="290" w:charSpace="{cs}"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(CS))


def _word_adv(pdf):
    import fitz
    pg = fitz.open(pdf)[0]; out = []
    for b in pg.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            cs = [c for sp in l["spans"] for c in sp["chars"]]
            if len(cs) >= 3 and cs[0]["c"] == "あ":
                out.append((round(l["spans"][0]["size"], 1), round(cs[1]["bbox"][0] - cs[0]["bbox"][0], 2)))
    return out


def _oxi_adv(exe, docx):
    import subprocess, tempfile
    t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + out, "--pages=1"], capture_output=True)
    g = json.load(open(out, encoding="utf-8")); pages = g["pages"] if isinstance(g, dict) else g
    gl = pages[0]["glyphs"] if isinstance(pages[0], dict) and "glyphs" in pages[0] else pages[0]
    rows = {}
    for x in gl:
        rows.setdefault(round(x.get("baseline", x.get("top", 0)), 0), []).append(x)
    res = []
    for y in sorted(rows):
        r = sorted(rows[y], key=lambda x: x["x"])
        if len(r) >= 2 and r[0]["char"] == "あ":
            res.append((r[0].get("font_size"), round(r[1]["x"] - r[0]["x"], 2)))
    return res


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in CS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in CS:
        res[arm] = _word_adv(OUT / f"{arm}.pdf")
        print(f"{arm:6s} charSpace {CS[arm]:6d} word (size, advance):", res[arm], flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in CS:
        print(f"{arm:6s} oxi:", _oxi_adv(os.path.abspath(exe), OUT / f"{arm}.docx"), " word:", word.get(arm), flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
