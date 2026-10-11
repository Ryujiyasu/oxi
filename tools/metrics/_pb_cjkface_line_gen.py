# -*- coding: utf-8 -*-
"""What sets the auto line height of an HG CJK face (HGP創英ﾌﾟﾚｾﾞﾝｽEB)?

ja/reference__5593c6a7: a compatLnSpc text box of HGP創英ﾌﾟﾚｾﾞﾝｽEB bold+outline
paragraphs steps 31.15pt at 22pt and 36.72 at 26pt in Word's PDF (1.412em),
while the face's win box is 1.0em (x 83/64 = 1.297em, the MS-face natural).

Arms: <where><face>_<props>   where: c = compatLnSpc box, b = body
      face: hgp / msg (MS Gothic); props: n (none), b (bold), o (outline), bo
Three one-line paragraphs «壱» «弐» «参» at 22pt, single auto spacing, after 0.

    python tools/metrics/_pb_cjkface_line_gen.py gen
    python tools/metrics/_pb_cjkface_line_gen.py word
    python tools/metrics/_pb_cjkface_line_gen.py oxi <renderer.exe>
"""
import json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_cjkface_line"
FACES = {"hgp": "HGP創英ﾌﾟﾚｾﾞﾝｽEB", "msg": "ＭＳ ゴシック"}
PROPS = {"n": "", "b": "<w:b/>", "o": "<w:outline/>", "bo": "<w:b/><w:outline/>"}
ARMS = {f"{w}{f}_{p}": (w, f, p) for w in ("c", "b") for f in FACES for p in PROPS}
TOKENS = ("壱", "弐", "参")


def paras(face, props):
    f = FACES[face]
    rpr = (f'<w:rPr><w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}" w:hint="eastAsia"/>{PROPS[props]}'
           '<w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr>')
    ppr = f'<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{rpr}</w:pPr>'
    return "".join(f"<w:p>{ppr}<w:r>{rpr}<w:t>{t}</w:t></w:r></w:p>" for t in TOKENS)


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (where, face, props) in ARMS.items():
        if where == "c":
            a = T.anchor(160, 1, False).replace(' anchor="t"', ' anchor="t" compatLnSpc="1"')
            a = re.sub(r"<w:txbxContent>.*</w:txbxContent>",
                       "<w:txbxContent>" + paras(face, props) + "</w:txbxContent>", a, flags=re.S)
            body = "<w:p>" + a + "</w:p>" + "".join(f"<w:p><w:r><w:t>L{i:02d}</w:t></w:r></w:p>" for i in range(1, 12))
        else:
            body = paras(face, props)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
               ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="425"/>'
               '<w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
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
        ys, fonts = {}, set()
        for b in fitz.open(OUT / f"{arm}.pdf")[0].get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        if c["c"] in TOKENS and c["c"] not in ys:
                            ys[c["c"]] = round(c["origin"][1], 3); fonts.add(s["font"])
        res[arm] = ys
        print(f"{arm:10s} first {ys.get(TOKENS[0])}  steps {_steps(ys)}  pdf font {sorted(fonts)}", flush=True)
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
        print(f"{arm:10s} first dy {d0}  oxi steps {_steps(ys)}  word {_steps(w)}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
