# -*- coding: utf-8 -*-
"""Which face draws a symbol the run's font lacks, and how much does it grow the
line?  blind-G ja/forms__02281020: «【□の中に✔を記入してください。】» in BIZ UD明朝
Medium 12pt; Word's PDF draws the ✔ (U+2714) in SegoeUIEmoji and the line's
first baseline sits 12.95 below the cell top (= Segoe UI's 1.079 em ascent, not
BIZ UD's 12.34), the row 0.6 taller; Oxi keeps the whole line in the base face
and the row 0.6 short (two rows, a -1.2 step for the rest of the page).
Arms: base face x symbol; three paragraphs «あいう» / «あ<sym>い» / «あいう»
(or «abc» for Latin bases) at 12pt, no grid, single spacing, after=0:
    bases   bizud (BIZ UD明朝 Medium), msmin (ＭＳ 明朝), msgo (ＭＳ ゴシック),
            yugo (游ゴシック), meiryo (Meiryo), calibri (Calibri, Latin text)
    syms    chk ✔ U+2714, chk2 ✓ U+2713, ballot ☑ U+2611, star ★ U+2605, sq □ U+25A1
Read: the PDF face of the symbol span and the three baselines (pitch 1->2, 2->3).

    python tools/metrics/_pb_symbol_fallback_face_gen.py gen
    python tools/metrics/_pb_symbol_fallback_face_gen.py word
    python tools/metrics/_pb_symbol_fallback_face_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_symbol_fallback_face"
BASES = {"bizud": "BIZ UD明朝 Medium", "msmin": "ＭＳ 明朝", "msgo": "ＭＳ ゴシック",
         "yugo": "游ゴシック", "meiryo": "Meiryo", "calibri": "Calibri",
         "yumin": "游明朝", "arial": "Arial", "tnr": "Times New Roman"}
SYMS = {"chk": "✔", "chk2": "✓", "ballot": "☑", "star": "★", "sq": "□",
        "box": "☐", "emdash": "—", "rquo": "’", "nbhy": "‑", "hy2": "‐", "nnbsp": " "}
ARMS = {b + "_" + s: (bf, sc) for b, bf in BASES.items() for s, sc in SYMS.items()}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (face, sym) in ARMS.items():
        latin = face in ("Calibri", "Arial", "Times New Roman")
        plain = "abc" if latin else "あいう"
        mid = ("a" + sym + "b") if latin else ("あ" + sym + "い")
        fonts = '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>' % (face, face, face, face)
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  + fonts + '<w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
                  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        def p(t):
            return '<w:p><w:r><w:t>' + t + '</w:t></w:r></w:p>'
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + p(plain) + p(mid) + p(plain) +
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
    for arm, (face, sym) in ARMS.items():
        pg = fitz.open(str(OUT / (arm + ".pdf")))[0]; ys = []; symfont = None
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip():
                        ys.append(round(s["origin"][1], 2))
                        if sym in s["text"]: symfont = s["font"] + (" adv=%.2f" % (s["bbox"][2] - s["bbox"][0]) if s["text"].strip() == sym else "")
        ys = sorted(set(ys)); p12 = ys[1] - ys[0] if len(ys) > 1 else None; p23 = ys[2] - ys[1] if len(ys) > 2 else None
        res[arm] = (symfont, ys)
        print("%-15s word sym=%-24s y=%s  p12=%s p23=%s" % (arm, symfont, ys, None if p12 is None else round(p12, 2), None if p23 is None else round(p23, 2)))
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    for arm, (face, sym) in ARMS.items():
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / (arm + ".docx")), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]
        ys = sorted({round(x["baseline"], 2) for x in g if x["char"].strip()})
        symfont = next((x["font_family"] for x in g if x["char"] == sym), None)
        p12 = ys[1] - ys[0] if len(ys) > 1 else None; p23 = ys[2] - ys[1] if len(ys) > 2 else None
        print("%-15s oxi  sym=%-24s y=%s  p12=%s p23=%s" % (arm, symfont, ys, None if p12 is None else round(p12, 2), None if p23 is None else round(p23, 2)))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
