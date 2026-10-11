# -*- coding: utf-8 -*-
"""Does a linesAndChars charSpace reach the text inside a text box?

ja/correspondence__14012ac2: a wps text box under <w:docGrid type="linesAndChars"
charSpace="-3930"> sets MS PGothic 12pt kanji 11.04pt apart (= 12 - 3930/4096)
and 10pt ones 9.04 apart in Word's PDF; Oxi lays text boxes out with no char
pitch (12 / 10) and wraps the heading one word early.

Arms: <where>_<face><size>_cs<charSpace>; where b = body, t = text box.
Each paragraph is «支所支部大会からの参加が可能となる競技» (+ «ⅢＡ・ら» probes).

    python tools/metrics/_pb_txbx_chargrid_gen.py gen
    python tools/metrics/_pb_txbx_chargrid_gen.py word
    python tools/metrics/_pb_txbx_chargrid_gen.py oxi <renderer.exe>
"""
import json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_txbx_chargrid"
FACES = {"pg": "ＭＳ Ｐゴシック", "g": "ＭＳ ゴシック"}
TEXT = "支所・支部大会からの参加が可能"
ARMS = {f"{w}_{f}{s}_cs{cs}": (w, f, s, cs)
        for w in ("b", "t") for f in FACES for s in (12, 10) for cs in (-3930, 1966, 0)}


def para(face, size):
    f = FACES[face]
    rpr = (f'<w:rPr><w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}" w:hint="eastAsia"/>'
           f'<w:sz w:val="{size * 2}"/></w:rPr>')
    return f'<w:p><w:pPr><w:jc w:val="left"/>{rpr}</w:pPr><w:r>{rpr}<w:t>{TEXT}</w:t></w:r></w:p>'


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (where, face, size, cs) in ARMS.items():
        if where == "t":
            a = T.anchor(160, 1, False)
            a = re.sub(r"<w:txbxContent>.*</w:txbxContent>", "<w:txbxContent>" + para(face, size) + "</w:txbxContent>",
                       a, flags=re.S)
            body = "<w:p>" + a + "</w:p>" + "".join(f"<w:p><w:r><w:t>L{i:02d}</w:t></w:r></w:p>" for i in range(1, 12))
        else:
            body = para(face, size)
        grid = (f'<w:docGrid w:type="linesAndChars" w:linePitch="398" w:charSpace="{cs}"/>' if cs
                else '<w:docGrid w:type="linesAndChars" w:linePitch="398"/>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
               ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="425"/>'
               + grid + '</w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", T.CT); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", T.DRELS); z.writestr("word/styles.xml", T.STYLES)
            z.writestr("word/document.xml", doc)
    print("ok", len(ARMS), "arms")


def advances(xs):
    return [round(xs[i + 1] - xs[i], 3) for i in range(len(xs) - 1)]


def word():
    import win32com.client, fitz
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in ARMS:
        xs = []
        for b in fitz.open(OUT / f"{arm}.pdf")[0].get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    t = "".join(c["c"] for c in s["chars"])
                    if TEXT[:3] in t:
                        xs = [c["origin"][0] for c in s["chars"]]
        res[arm] = advances(xs)
        print(f"{arm:14s} {res[arm][:8]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


def oxi(exe):
    import subprocess
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        t = tempfile.mkdtemp(); out = os.path.join(t, "l.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-layout=" + out],
                       capture_output=True, timeout=300)
        els = [e for e in json.load(open(out, encoding="utf-8"))["pages"][0]["elements"]
               if e["type"] == "text" and e["text"] in TEXT and e["text"]]
        xs = []
        for e in sorted(els, key=lambda e: e["x"]):
            if len(e["text"]) == 1:
                xs.append(e["x"])
        print(f"{arm:14s} oxi {advances(xs)[:6]}  word {word[arm][:6]}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
