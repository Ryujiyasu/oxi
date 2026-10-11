# -*- coding: utf-8 -*-
"""Does the compatLnSpc text-box line law (CLNSP: auto line = 1.27fs + 0.25,
no line multiple, no space-after) hold for LATIN text?

en/policies__009ab5c3: nine compatLnSpc="1" wps boxes of Arial 14 paragraphs
(docDefaults line 259, after 160). Word's PDF steps 17.40 inside a paragraph
and 25.32 across one (= 17.40 + 8); Oxi's CLNSP chain gives 18.03 and 18.03.
The law was derived on MS Mincho / MS Gothic boxes (_compatlnspc_sweep.py).

One floating box per arm (the _pb_txbxfit_gen.py anchor), three one-line
paragraphs «Qa» «Wb» «Ec» with direct spacing line/after; read Word's PDF
baselines -> per-paragraph step.

Arm name: <c|n><face><size>_<line>_<after>   c = bodyPr compatLnSpc="1"

    python tools/metrics/_pb_clnsp_latin_gen.py gen
    python tools/metrics/_pb_clnsp_latin_gen.py word
    python tools/metrics/_pb_clnsp_latin_gen.py oxi <renderer.exe>
"""
import json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402  (anchor / NS / package parts)

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_clnsp_latin"
FACES = {"ar": "Arial", "cal": "Calibri", "tnr": "Times New Roman", "msg": "ＭＳ ゴシック"}
# (compat, face key, half-point size, line (240ths, auto), after twips)
ARMS = {}
for c in (True, False):
    for face, hsz, line, after in (("ar", 28, 259, 160), ("ar", 28, 240, 0), ("ar", 28, 259, 0),
                                   ("ar", 28, 240, 160), ("cal", 22, 259, 160), ("tnr", 24, 240, 0),
                                   ("msg", 24, 240, 0), ("msg", 24, 259, 160)):
        ARMS[f"{'c' if c else 'n'}{face}{hsz // 2}_{line}_{after}"] = (c, face, hsz, line, after)
# the real doc is compatibilityMode 14: the same arms with a declared mode
for k in (14, 15):
    for name in ("car14_259_160", "nar14_259_160", "cmsg12_259_160", "nmsg12_259_160"):
        ARMS[f"{name}_k{k}"] = ARMS[name] + (k,)
# typed docGrid (lines, 18pt): does compatLnSpc only switch the box off the grid?
for name in ("car14_259_160", "nar14_259_160", "cmsg12_240_0", "nmsg12_240_0", "cmsg12_259_160", "nmsg12_259_160"):
    ARMS[f"{name}_g"] = ARMS[name][:5] + (14, "grid")
TOKENS = ("Qa", "Wb", "Ec")
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>')


# exact / atLeast lines (the 13:3 split / atLeast floor of the CLNSP law)
for c in (True, False):
    for face, hsz, rule, val in (("msg", 24, "exact", 480), ("msg", 24, "atLeast", 340), ("ar", 28, "exact", 480),
                                 ("ar", 28, "atLeast", 340)):
        ARMS[f"{'c' if c else 'n'}{face}{hsz // 2}_{rule}{val}"] = (c, face, hsz, (rule, val), 0)


def paras(face, hsz, line, after):
    f = FACES[face]
    rpr = (f'<w:rPr><w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}" w:cs="{f}"/>'
           f'<w:sz w:val="{hsz}"/><w:szCs w:val="{hsz}"/></w:rPr>')
    rule = "auto"
    if isinstance(line, tuple):
        rule, line = line
    ppr = (f'<w:pPr><w:spacing w:before="0" w:after="{after}" w:line="{line}" w:lineRule="{rule}"/>'
           f'{rpr}</w:pPr>')
    return "".join(f"<w:p>{ppr}<w:r>{rpr}<w:t>{t}</w:t></w:r></w:p>" for t in TOKENS)


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (compat, face, hsz, line, after, *mode) in ARMS.items():
        a = T.anchor(160, 1, False)
        if compat:
            a = a.replace(' anchor="t"', ' anchor="t" compatLnSpc="1"')
        a = re.sub(r"<w:txbxContent>.*</w:txbxContent>",
                   "<w:txbxContent>" + paras(face, hsz, line, after) + "</w:txbxContent>", a, flags=re.S)
        body = "<w:p>" + a + "</w:p>" + "".join(f"<w:p><w:r><w:t>L{i:02d}</w:t></w:r></w:p>" for i in range(1, 12))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
               ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="425"/>'
               '<w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            ct = T.CT if not mode else T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType='
                                                    '"application/vnd.openxmlformats-officedocument.wordprocessingml.'
                                                    'settings+xml"/></Types>')
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            drels = T.DRELS
            if mode:
                drels = drels.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                                      'officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
                z.writestr("word/settings.xml", SETTINGS % mode[0])
            if mode[1:] == ["grid"]:
                doc = doc.replace('<w:docGrid w:linePitch="360"/>', '<w:docGrid w:type="lines" w:linePitch="360"/>')
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", T.STYLES)
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
                    txt = "".join(c["c"] for c in s["chars"])
                    for t in TOKENS:
                        if t in txt and t not in ys:
                            ys[t] = round(s["chars"][txt.index(t)]["origin"][1], 3)
        res[arm] = ys
        print(f"{arm:18s} first {ys.get(TOKENS[0])}  steps {_steps(ys)}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


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
        print(f"{arm:18s} first dy {d0}  oxi steps {_steps(ys)}  word {_steps(w)}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
