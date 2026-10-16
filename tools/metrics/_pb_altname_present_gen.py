# -*- coding: utf-8 -*-
"""Does Word honour a fontTable altName when the PRIMARY face is available?

blind-H en/technical__01838a08 declares «Roboto Condensed» with
<w:altName w:val="Arial"/>; the face is not in Windows\\Fonts but sits in the
Office cloud cache, and Word's PDF draws the heading in Roboto Condensed Bold.
Oxi took the altName (the compiled tables lack the family) and drew Arial: the
title 18pt off in x and the line box 0.7 short.

Arms <face>_<alt>: face = a cloud-cache face (Roboto Condensed, Lato, Roboto),
a system face missing from the compiled tables (Candara, Corbel, Bahnschrift),
and a face that exists nowhere (Nowhere Sans); alt = arial / none.  One 18pt
bold paragraph «Heading Probe» and an 11pt paragraph «Body probe text».
Read: the PDF font name per paragraph and the baselines (Word; Oxi glyphs).

    python tools/metrics/_pb_altname_present_gen.py gen
    python tools/metrics/_pb_altname_present_gen.py word
    python tools/metrics/_pb_altname_present_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_altname_present"
FACES = {"robc": "Roboto Condensed", "lato": "Lato", "rob": "Roboto", "cand": "Candara", "corb": "Corbel",
         "bahn": "Bahnschrift", "nowhere": "Nowhere Sans"}
ARMS = {f"{k}_{a}": (k, a) for k in FACES for a in ("arial", "none")}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (k, a) in ARMS.items():
        face = FACES[k]
        rf = f'<w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:cs="{face}"/>'
        body = (f'<w:p><w:r><w:rPr>{rf}<w:b/><w:sz w:val="36"/></w:rPr><w:t>Heading Probe</w:t></w:r></w:p>'
                f'<w:p><w:r><w:rPr>{rf}<w:sz w:val="22"/></w:rPr><w:t>Body probe text</w:t></w:r></w:p>'
                '<w:p><w:r><w:t>Tail line</w:t></w:r></w:p>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
               '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
               '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
        alt = '<w:altName w:val="Arial"/>' if a == "arial" else ''
        fonts = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts ' + T.NS + '>'
                 f'<w:font w:name="{face}">{alt}<w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>'
                 '<w:font w:name="Calibri"><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font></w:fonts>')
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:lang w:val="en-US"/>'
                  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" '
                    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                          '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                                '<Relationship Id="rIdFt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings); z.writestr("word/fontTable.xml", fonts); z.writestr("word/document.xml", doc)
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
        pg = fitz.open(str(OUT / f"{arm}.pdf"))[0]; r = {}
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                s = l["spans"][0]; t = "".join(x["text"] for x in l["spans"])
                key = "head" if t.startswith("Heading") else "body" if t.startswith("Body") else "tail"
                r[key] = (round(s["origin"][1], 2), s["font"], round(l["bbox"][2] - l["bbox"][0], 1))
        res[arm] = r
        print(f"{arm:14s} word {r}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]; rows = {}
        for x in g:
            rows.setdefault(round(x["baseline"], 2), []).append(x)
        r = {}
        for y, xs in sorted(rows.items()):
            t = "".join(x["char"] for x in xs)
            key = "head" if t.startswith("Heading") else "body" if t.startswith("Body") else "tail"
            r[key] = (y, xs[0]["font_family"], round(xs[-1]["x"] - xs[0]["x"], 1))
        w = wd[arm]
        flag = "" if r.get("tail") and w.get("tail") and abs(r["tail"][0] - w["tail"][0]) < 0.3 else "  <-- tail differs"
        print(f"{arm:14s} oxi {r}\n{'':14s} word {w}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
