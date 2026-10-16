# -*- coding: utf-8 -*-
"""How much room does a COMPOUND paragraph border take, and how is it drawn?

blind-H en/correspondence__00c28966: two consecutive centred paragraphs share
<w:bottom w:val="thinThickSmallGap" w:sz="12" w:space="1"/>.  Word draws two
bands under the group (thick 103.58-105.02, gap, thin 105.74-106.46 = 2.88pt
in all for sz 12 = 1.5pt) and the body that follows sits 1.5 lower than Oxi's,
whose dump carries no rule at all.  Only double / triple are measured so far
(`_pb_tblborder_gen.py`: 3x and 5x sz/8 for cell borders).

Arms <style>_<sz>: one Calibri 11 paragraph «Bordered» with a bottom border of
that style and size (space 1), followed by «After line»; style = none / single
/ double / triple / thinThickSmallGap / thickThinSmallGap / thinThickMediumGap
/ thickThinMediumGap / thinThickLargeGap / thickThinLargeGap /
thinThickThinSmallGap / thinThickThinMediumGap / dashDotStroked / threeDEmboss
/ threeDEngrave / thinThickThinLargeGap / wave / doubleWave; sz 4 / 12 / 24.
Read: the «After line» baseline minus the «none» arm's, and the drawn bands
under «Bordered» (Word PDF drawings; Oxi dump-layout border elements).

    python tools/metrics/_pb_pbdr_compound_gen.py gen
    python tools/metrics/_pb_pbdr_compound_gen.py word
    python tools/metrics/_pb_pbdr_compound_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_pbdr_compound"
STYLES_LIST = ["single", "double", "triple", "thinThickSmallGap", "thickThinSmallGap", "thinThickMediumGap",
               "thickThinMediumGap", "thinThickLargeGap", "thickThinLargeGap", "thinThickThinSmallGap",
               "thinThickThinMediumGap", "thinThickThinLargeGap", "dashDotStroked", "threeDEmboss", "threeDEngrave",
               "wave", "doubleWave"]
ARMS = {"none_0": ("none", 0)}
for st in STYLES_LIST:
    for sz in (4, 12, 24):
        ARMS[f"{st}_{sz}"] = (st, sz)

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (st, sz) in ARMS.items():
        bdr = "" if st == "none" else f'<w:pBdr><w:bottom w:val="{st}" w:sz="{sz}" w:space="1" w:color="943634"/></w:pBdr>'
        body = (f'<w:p><w:r><w:t>Top line</w:t></w:r></w:p><w:p><w:pPr>{bdr}</w:pPr><w:r><w:t>Bordered</w:t></w:r></w:p>'
                '<w:p><w:r><w:t>After line</w:t></w:r></w:p>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
               '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
               '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" '
                    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


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
        pg = fitz.open(str(OUT / f"{arm}.pdf"))[0]; ys = {}
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).strip()
                ys[t.split()[0]] = round(l["spans"][0]["origin"][1], 2)
        bands = sorted({(round(d["rect"].y0, 2), round(d["rect"].y1, 2)) for d in pg.get_drawings() if d["rect"].width > 50})
        res[arm] = {"bordered": ys.get("Bordered"), "after": ys.get("After"), "bands": bands}
    base = res["none_0"]["after"]
    for arm, r in res.items():
        r["adv"] = round(r["after"] - base, 2) if r["after"] and base else None
        print(f"{arm:28s} word adv {r['adv']}  bands {r['bands']}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    res = {}
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json"); lo = os.path.join(t, "l.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "q"), "--dump-layout=" + lo], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]
        rows = {}
        for x in g:
            rows.setdefault(round(x["baseline"], 2), []).append(x["char"])
        after = [y for y, c in rows.items() if "".join(c).startswith("After")]
        L = json.load(open(lo, encoding="utf-8")); pg = L["pages"][0] if isinstance(L, dict) and "pages" in L else L[0]
        els = pg.get("elements", pg) if isinstance(pg, dict) else pg
        bands = sorted({(round(e["y"], 2), round(e["y"] + e["h"], 2)) for e in els if e.get("type") in ("border", "shading", "line", "rect") and e.get("w", 0) > 50 and e.get("h", 0) < 10})
        res[arm] = {"after": after[0] if after else None, "bands": bands}
    base = res["none_0"]["after"]
    for arm, r in res.items():
        adv = round(r["after"] - base, 2) if r["after"] and base else None
        w = wd[arm]
        flag = "" if adv is not None and w["adv"] is not None and abs(adv - w["adv"]) < 0.3 else "  <-- differs"
        print(f"{arm:28s} oxi adv {adv}  word adv {w['adv']}{flag}   oxi bands {r['bands']}  word {w['bands']}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
