# -*- coding: utf-8 -*-
"""Baseline of an exact / atLeast line that mixes a Latin face with a CJK
fallback glyph.  blind-G en/creative__03a7fb6a (Times New Roman 12pt, line 480
atLeast, a full-width comma U+FF0C drawn in MS-Mincho): Word's mixed line sits
0.84 HIGHER than its plain neighbours (pitch 23.16 in, 24.84 out) and both
faces share the baseline; Oxi puts it 0.8 lower with the faces 0.9 apart.
Hypothesis: an exact/atLeast line is bottom-aligned, baseline = box bottom -
the line's max descent, and the ＭＳ 明朝 glyph's deeper box (12pt: 3.47 against
Times' 2.59) lifts it.  Arms (A4, Times New Roman 12pt via docDefaults, three
paragraphs plain / mixed / plain then «end»):
    al480_comma   atLeast 480, «water，they will»      al480_plain  control (no comma)
    ex480_comma   exact 480                           ex300_comma  exact 300
    au240_comma   auto 240 (single)                   al480_kanji  atLeast 480, «water水they»
    al480_ms      atLeast 480, comma with eastAsia ＭＳ 明朝 declared
Read: the three baselines and the faces (Word PDF; Oxi dump).

    python tools/metrics/_pb_exact_mixed_desc_gen.py gen
    python tools/metrics/_pb_exact_mixed_desc_gen.py word
    python tools/metrics/_pb_exact_mixed_desc_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_exact_mixed_desc"
TNR = "Times New Roman"
MS = "ＭＳ 明朝"
# (rule, line twips, middle text, eastAsia face)
ARMS = {"al480_comma": ("atLeast", 480, "water，they will be able", TNR), "al480_plain": ("atLeast", 480, "water, they will be able", TNR),
        "ex480_comma": ("exact", 480, "water，they will be able", TNR), "ex300_comma": ("exact", 300, "water，they will be able", TNR),
        "au240_comma": ("auto", 240, "water，they will be able", TNR), "al480_kanji": ("atLeast", 480, "water水they will be able", TNR),
        "al480_ms": ("atLeast", 480, "water，they will be able", MS)}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (rule, line, mid, ea) in ARMS.items():
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:rFonts w:ascii="' + TNR + '" w:hAnsi="' + TNR + '" w:eastAsia="' + ea + '" w:cs="' + TNR + '"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US"/>'
                  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="%d" w:lineRule="%s"/></w:pPr></w:pPrDefault></w:docDefaults>' % (line, rule) +
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        p = lambda t: '<w:p><w:r><w:t xml:space="preserve">' + t + '</w:t></w:r></w:p>'
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + p("Access to clean water is a stepping") + p(mid) + p("water is becoming more difficult") + p("end") +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / (arm + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


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
        pg = fitz.open(str(OUT / (arm + ".pdf")))[0]; rows = {}
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip(): rows.setdefault(round(s["origin"][1], 2), set()).add(s["font"])
        ys = sorted(rows); res[arm] = [(y, sorted(rows[y])) for y in ys]
        print("%-12s word ys=%s pitches=%s faces=%s" % (arm, ys, [round(ys[i + 1] - ys[i], 2) for i in range(len(ys) - 1)], [sorted(rows[y]) for y in ys][1]))
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / (arm + ".docx")), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]; rows = {}
        for x in g:
            if x["char"].strip(): rows.setdefault(round(x["baseline"], 2), set()).add(x["font_family"])
        ys = sorted(rows)
        print("%-12s oxi  ys=%s pitches=%s faces=%s" % (arm, ys, [round(ys[i + 1] - ys[i], 2) for i in range(len(ys) - 1)], [sorted(rows[y]) for y in ys]))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
