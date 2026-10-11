# -*- coding: utf-8 -*-
"""Where does Word put the baseline of a CJK line that takes several cells of a
typed (lines) docGrid?

ja/administrative__199dab4d: a centred 22pt title (theme minorEastAsia = 游明朝)
on a 18pt lines grid takes two cells; Word's PDF baseline sits 3.08pt above
Oxi's while the following lines agree (the box is right, the baseline is not).

Arms: <face>_<size>: the first paragraph «永字八法» at the top margin, then a
10.5pt «参照» line. Read the first baseline's offset from the top margin.

    python tools/metrics/_pb_gridline_base_gen.py gen
    python tools/metrics/_pb_gridline_base_gen.py word
    python tools/metrics/_pb_gridline_base_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_gridline_base"
FACES = {"yum": "游明朝", "msm": "ＭＳ 明朝", "yug": "游ゴシック", "msg": "ＭＳ ゴシック"}
SIZES = (12, 16, 20, 22, 28)
ARMS = {f"{f}_{s}": (f, s) for f in FACES for s in SIZES}
TOP = 1440 / 20.0
TOKEN = "永字八法"


def body(face, size):
    f = FACES[face]
    rpr = (f'<w:rPr><w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}" w:hint="eastAsia"/>'
           f'<w:sz w:val="{size * 2}"/></w:rPr>')
    rpr2 = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="21"/></w:rPr>'
    return (f'<w:p><w:pPr><w:jc w:val="center"/>{rpr}</w:pPr><w:r>{rpr}<w:t>{TOKEN}</w:t></w:r></w:p>'
            f'<w:p><w:pPr>{rpr2}</w:pPr><w:r>{rpr2}<w:t>参照</w:t></w:r></w:p>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (face, size) in ARMS.items():
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" +
               body(face, size) +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
               ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="425"/>'
               '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", T.CT); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", T.DRELS); z.writestr("word/styles.xml", T.STYLES)
            z.writestr("word/document.xml", doc)
    print("ok", len(ARMS), "arms")


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
        ys = {}
        for b in fitz.open(OUT / f"{arm}.pdf")[0].get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        if c["c"] in "永参" and c["c"] not in ys:
                            ys[c["c"]] = round(c["origin"][1], 3)
        res[arm] = ys
        print(f"{arm:8s} base-top {round(ys.get('永', 0) - TOP, 3)}  next {round(ys.get('参', 0) - TOP, 3)}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        ys = {}
        for o in L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"):
            for ch in "永参":
                if o["t"].startswith(ch) and ch not in ys:
                    ys[ch] = o["y"]
        w = word[arm]
        print(f"{arm:8s} dy first {round(ys.get('永', 0) - w.get('永', 0), 2)}  next {round(ys.get('参', 0) - w.get('参', 0), 2)}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
