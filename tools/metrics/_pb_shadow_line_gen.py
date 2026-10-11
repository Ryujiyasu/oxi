# -*- coding: utf-8 -*-
"""How much does a w14:shadow text effect add to a line's height?

ja/reference__5593c6a7: HGP創英ﾌﾟﾚｾﾞﾝｽEB lines carrying
<w14:shadow blurRad=50800 dist=38100 dir=5400000> step 31.11pt at 22pt and
36.72 at 26pt in Word's PDF; with the shadow removed (XML strip) 28.56 = the
plain 83/64 line. Oxi ignores the effect.

Arms (body paragraphs, single auto spacing, three one-line paragraphs):
  <face><size>_<shadow>   face: msg (MS Gothic) / ar (Arial); shadow keys below.

    python tools/metrics/_pb_shadow_line_gen.py gen
    python tools/metrics/_pb_shadow_line_gen.py word
    python tools/metrics/_pb_shadow_line_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_shadow_line"
W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
FACES = {"msg": "ＭＳ ゴシック", "ar": "Arial"}
# (blurRad EMU, dist EMU, dir 60000ths deg)
SHADOWS = {"none": None, "A": (50800, 38100, 5400000), "B": (38100, 22860, 5400000),
           "d0": (50800, 0, 5400000), "b0": (0, 38100, 5400000), "right": (50800, 38100, 0),
           "up": (50800, 38100, 16200000), "d6": (50800, 76200, 5400000), "b8": (101600, 38100, 5400000)}
ARMS = {}
for face, size, keys in (("msg", 22, SHADOWS), ("msg", 12, ("none", "A")), ("msg", 26, ("none", "A")),
                         ("ar", 22, ("none", "A", "B")),
                         # the size law: T = 3 / 8 / 16 at 8..40pt
                         ("msg", 8, ("none", "b0", "A", "b8")), ("msg", 16, ("none", "b0", "A", "b8")),
                         ("msg", 30, ("none", "b0", "A", "b8")), ("msg", 40, ("none", "b0", "A", "b8")),
                         ("msg", 12, ("b0", "b8")), ("msg", 26, ("b0", "b8"))):
    for k in keys:
        ARMS[f"{face}{size}_{k}"] = (face, size, k)
TOKENS = ("壱", "弐", "参")


def shadow_xml(k):
    if SHADOWS[k] is None:
        return ""
    b, d, a = SHADOWS[k]
    return (f'<w14:shadow w14:blurRad="{b}" w14:dist="{d}" w14:dir="{a}" w14:sx="100000" w14:sy="100000" '
            'w14:kx="0" w14:ky="0" w14:algn="t"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="60000"/>'
            '</w14:srgbClr></w14:shadow>')


def paras(face, size, k):
    f = FACES[face]
    rpr = (f'<w:rPr><w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}" w:hint="eastAsia"/>'
           f'<w:sz w:val="{size * 2}"/><w:szCs w:val="{size * 2}"/>{shadow_xml(k)}</w:rPr>')
    ppr = f'<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{rpr}</w:pPr>'
    return "".join(f"<w:p>{ppr}<w:r>{rpr}<w:t>{t}</w:t></w:r></w:p>" for t in TOKENS)


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    ns = T.NS if "xmlns:w14" in T.NS else T.NS + " " + W14
    for arm, (face, size, k) in ARMS.items():
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + ns +
               ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>' +
               paras(face, size, k) +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
               ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="425"/>'
               '<w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
        doc = doc.replace(' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"',
                          "", 1) if "xmlns:mc=" in ns else doc
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", T.CT); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", T.DRELS); z.writestr("word/styles.xml", T.STYLES)
            z.writestr("word/document.xml", doc)
    print("ok", len(ARMS), "arms")


def _steps(ys):
    return [round(ys[TOKENS[i + 1]] - ys[TOKENS[i]], 3) for i in range(2)] if all(t in ys for t in TOKENS) else None


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
                        if c["c"] in TOKENS and c["c"] not in ys:
                            ys[c["c"]] = round(c["origin"][1], 3)
        res[arm] = ys
        print(f"{arm:12s} first {ys.get(TOKENS[0])}  steps {_steps(ys)}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8")) if (OUT / "word.json").exists() else {}
    for arm in ARMS:
        ys = {}
        for o in L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"):
            for t in TOKENS:
                if o["t"].startswith(t) and t not in ys:
                    ys[t] = round(o["y"], 3)
        w = word.get(arm, {})
        d0 = round(ys[TOKENS[0]] - w[TOKENS[0]], 2) if TOKENS[0] in ys and TOKENS[0] in w else None
        print(f"{arm:12s} first dy {d0}  oxi steps {_steps(ys)}  word {_steps(w)}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
