# -*- coding: utf-8 -*-
"""What paragraph spacing does Word apply when styles.xml has NO pPrDefault?

en/reports__00ac5c81 (compat 12, docDefaults rPr only, empty Normal, jc=both
Calibri 12): Word's PDF steps 16.95 inside a paragraph and 24.96 across one
(= 16.95 + 8.0); Oxi uses the single 14.65 line and no after.

Arms: <compat>_<styles>: compat none / 12 / 14 / 15; styles = noppr (docDefaults
without pPrDefault), empty (<w:pPrDefault/>), zero (spacing after=0 line=240).
Four one-line paragraphs P1..P4, Calibri 12; read the PDF baseline steps.

    python tools/metrics/_pb_implicit_ppr_gen.py gen
    python tools/metrics/_pb_implicit_ppr_gen.py word
    python tools/metrics/_pb_implicit_ppr_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_implicit_ppr"
PPR = {
    "noppr": "",
    "empty": "<w:pPrDefault/>",
    "zero": '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>',
}
COMPAT = {"none": None, "c12": 12, "c14": 14, "c15": 15}
ARMS = {f"{c}_{p}": (c, p) for c in COMPAT for p in PPR}
# size law: no pPrDefault at 10 / 20pt, and a WRAPPED paragraph to split line pitch from after
SIZES = {"none_noppr": 24, "sz10_noppr": 20, "sz20_noppr": 40, "wrap_noppr": 24}
for a in ("sz10_noppr", "sz20_noppr", "wrap_noppr"):
    ARMS[a] = ("none", "noppr")
TOKENS = ("P1x", "P2x", "P3x", "P4x")


def styles(ppr, hsz=24):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults>'
            '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>' % (hsz, hsz) + PPR[ppr] +
            '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
            '</w:style></w:styles>')


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat>'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/>'
            '</w:compat></w:settings>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (c, p) in ARMS.items():
        compat = COMPAT[c]
        filler = " body text that runs on and on across the whole measure to force a wrap" * (3 if arm.startswith("wrap") else 0)
        paras = "".join(f'<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>{t} body text{filler}</w:t></w:r></w:p>' for t in TOKENS)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + paras +
               '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
               ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        ct, drels = T.CT, T.DRELS
        if compat:
            ct = ct.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.'
                            'openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
            drels = drels.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/'
                                  'officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles(p, SIZES.get(arm, 24)))
            if compat:
                z.writestr("word/settings.xml", settings(compat))
            z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def _steps(ys):
    return [round(ys[TOKENS[i + 1]] - ys[TOKENS[i]], 2) for i in range(3)] if all(t in ys for t in TOKENS) else None


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
        for b in fitz.open(OUT / f"{arm}.pdf")[0].get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                tx = "".join(s["text"] for s in l["spans"])
                for k in TOKENS:
                    if tx.startswith(k) and k not in ys:
                        ys[k] = round(l["spans"][0]["origin"][1], 3)
                if arm.startswith("wrap") and "P1x" in ys and "L2" not in ys and not any(tx.startswith(k) for k in TOKENS) and l["spans"][0]["origin"][1] > ys["P1x"] + 1:
                    ys["L2"] = round(l["spans"][0]["origin"][1], 3)
        res[arm] = ys
        if "L2" in ys: print(f"   {arm}: line pitch inside P1 = {round(ys['L2'] - ys['P1x'], 2)}")
        print(f"{arm:12s} first {ys.get(TOKENS[0])}  steps {_steps(ys)}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        ys = {}
        for o in L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"):
            for k in TOKENS:
                if o["t"].startswith(k) and k not in ys:
                    ys[k] = round(o["y"], 3)
        print(f"{arm:12s} oxi steps {_steps(ys)}  word {_steps(word[arm])}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
