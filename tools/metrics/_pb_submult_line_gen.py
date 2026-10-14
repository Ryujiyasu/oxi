# -*- coding: utf-8 -*-
"""Where does Word cut the line box when the spacing multiple is BELOW one
(w:line="220" auto = 0.9167)?  blind-G en/reports__009e23be (Skolar Sans and
Museo Slab missing -> Arial; title 17/18pt bold, body 11pt, every paragraph
line=220): Oxi's title sits 1.4pt lower than Word's and the body 0.8pt lower.
Arms (A4, margins 1440, Arial via docDefaults, explicit pPrDefault after=0,
three paragraphs «Hello world» x3 then a 12pt empty paragraph then «end»):
    m{220,200,180,240,276,300}_{11,18}     multiple x size (sz 22 / 36)
    mix220  line 220 with runs 17pt bold + 18pt bold (the title shape)
    emp220  line 220, 11pt body with an EMPTY 12pt paragraph between
Read: every baseline (Word PDF; Oxi dump): first baseline from the top margin
(72) tells the ascent part of the cut box; the pitch tells the box.

    python tools/metrics/_pb_submult_line_gen.py gen
    python tools/metrics/_pb_submult_line_gen.py word
    python tools/metrics/_pb_submult_line_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_submult_line"
ARMS = {}
for m in (220, 200, 180, 240, 276, 300):
    for sz in (22, 36):
        ARMS["m%d_%d" % (m, sz // 2)] = ("plain", m, sz)
ARMS["mix220"] = ("mix", 220, 36)
ARMS["emp220"] = ("emp", 220, 22)


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (kind, m, sz) in ARMS.items():
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US"/>'
                  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        sp = '<w:spacing w:line="%d" w:lineRule="auto"/>' % m

        def para(runs, extra=""):
            return '<w:p><w:pPr>' + sp + extra + '</w:pPr>' + runs + '</w:p>'
        run = lambda t, s, b=False: '<w:r><w:rPr>' + ('<w:b/>' if b else '') + '<w:sz w:val="%d"/></w:rPr><w:t xml:space="preserve">' % s + t + '</w:t></w:r>'
        if kind == "plain":
            body = "".join(para(run("Hello world %d" % i, sz)) for i in range(3))
        elif kind == "mix":
            body = para(run("FACT SHEET: ", 34, True) + run("Specialists here", 36, True)) + para(run("second line of the title", 36, True)) + para(run("third", 36, True))
        else:
            body = para(run("Hello world 0", 22)) + '<w:p><w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr></w:p>' + para(run("Hello world 1", 22)) + para(run("Hello world 2", 22))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
               '<w:p><w:r><w:t>end</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / (arm + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def fmt(arm, who, ys):
    p = ["%.2f" % (ys[i + 1] - ys[i]) for i in range(len(ys) - 1)]
    print("%-10s %s y0-72=%6.2f  ys=%s  pitches=%s" % (arm, who, ys[0] - 72.0, [round(y, 2) for y in ys], p))


def word():
    import win32com.client, fitz
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / (arm + ".docx"), t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / (arm + ".pdf")).resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in ARMS:
        pg = fitz.open(str(OUT / (arm + ".pdf")))[0]; ys = set()
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip(): ys.add(round(s["origin"][1], 2))
        ys = sorted(ys); res[arm] = ys; fmt(arm, "word", ys)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / (arm + ".docx")), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]
        ys = sorted({round(x["baseline"], 2) for x in g if x["char"].strip()})
        fmt(arm, "oxi ", ys)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
