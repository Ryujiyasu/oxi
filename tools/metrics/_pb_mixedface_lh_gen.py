# -*- coding: utf-8 -*-
"""Line height and baseline of a line whose CJK characters are drawn by several
faces (Word's per-character fallback under a Latin eastAsia face, or declared
faces).  blind-G ja/reports__597cb86c (docDefaults Calibri incl. eastAsia, no
pPrDefault, compat 12, Chinese body): Word's PDF draws JIS kanji in MS Mincho,
simplified-only characters in SimSun, 类 in Yu Gothic and 页 in Noto Sans JP
(labelled "Calibri" in the PDF, box 1.160/0.288); the three single-line items
«1. … / 2. … / 3. …» pitch 30.48 / 31.80 / 28.44 while a Mincho+SimSun line
inside a paragraph pitches 18.0 (= 12 × 83/64 × 278/240).  Oxi draws some
faces on their own baseline and pitches 26.5 / 34.2.
Arms (12pt Calibri incl. eastAsia unless declared, compat 12, no pPrDefault →
the endpoint's after=160 line=278; each arm = the SAME paragraph three times,
then «end»):
    item1 / item2 / item3   the document's three items
    mincho  JIS kanji only (→ MS Mincho on this machine)    simsun  simplified-only (→ SimSun)
    ye      页页页页 (→ Noto Sans JP)                         lei     类类类类 (→ Yu Gothic)
    latin   cookie cookie cookie                             lat_min cookie 当私偏好 cookie
    seq     items 1, 2, 3 once each
    *_th    the document's theme carried along (Oxi's fallback follows the theme)
    decl_*  the same faces DECLARED as the eastAsia font (no fallback involved)
    mix_split   Noto Sans JP 12pt (a 16.58 / d 6.06) beside Yu Mincho 13.5pt (a 16.09 / d 6.45):
                union of per-face ascent/descent → natural 23.03; tallest face (Noto 22.64) → 22.64
    mix_big_lat Calibri 16pt (a 15.23 / d 4.30) beside ＭＳ 明朝 12pt (12.14 / 3.42) — Latin governs both
    mix_noto_ms Noto Sans JP 12pt 页 beside ＭＳ 明朝 12pt
Read: every line's baseline y and the faces on it (Word PDF spans; Oxi --dump-glyphs).

    python tools/metrics/_pb_mixedface_lh_gen.py gen
    python tools/metrics/_pb_mixedface_lh_gen.py word
    python tools/metrics/_pb_mixedface_lh_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_mixedface_lh"
I1 = "1. 当访问隐私偏好中心网站时，会使用 cookie 存储和检索浏览器中的信息。"
I2 = "2. 这些 cookie 可以帮助个性化网页体验，但用户可以选择不使用某些 cookie。"
I3 = "3. 不同类型的 cookie 对网站功能、性能和广告定向等方面有不同影响。"
ARMS = {"item1": [I1] * 3, "item2": [I2] * 3, "item3": [I3] * 3,
        "mincho": ["当私偏好中心网站会使用存和索器中的信息。"] * 3, "simsun": ["访问隐时储检浏览"] * 3,
        "ye": ["页页页页"] * 3, "lei": ["类类类类"] * 3, "latin": ["cookie cookie cookie"] * 3,
        "lat_min": ["cookie 当私偏好 cookie"] * 3, "seq": [I1, I2, I3],
        "seq_th": [I1, I2, I3], "mincho_th": ["当私偏好中心网站会使用存和索器中的信息。"] * 3, "head_th": ["Article summary:", I1, I2],
        "decl_ms": ["当私偏好中心网站会使用存和索器中的信息。"] * 3, "decl_simsun": ["访问隐时储检浏览"] * 3,
        "decl_noto": ["页页页页"] * 3, "decl_yug": ["类类类类"] * 3, "decl_yum": ["当私偏好中心网站"] * 3,
        "mix_split": [[("Noto Sans JP", 24, "页页页"), ("Yu Mincho", 27, "当当当")]] * 3,
        "mix_big_lat": [[("Calibri", 32, "cookie "), ("ＭＳ 明朝", 24, "当当当")]] * 3,
        "mix_noto_ms": [[("Noto Sans JP", 24, "页"), ("ＭＳ 明朝", 24, "当当当当当")]] * 3}
DECL = {"decl_ms": "ＭＳ 明朝", "decl_simsun": "宋体", "decl_noto": "Noto Sans JP", "decl_yug": "Yu Gothic", "decl_yum": "Yu Mincho"}
THEME_SRC = REPO / "pipeline_data" / "docx_corpus" / "ja" / "reports" / "597cb86c6b6de30f.docx"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/>'
              '</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/></w:compat></w:settings>')
    ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                      '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>')
    drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    theme = zipfile.ZipFile(THEME_SRC).read("word/theme/theme1.xml")
    for arm, paras in ARMS.items():
        ea = DECL.get(arm, "Calibri")
        st = styles.replace('w:eastAsia="Calibri"', f'w:eastAsia="{ea}"')

        def para(p):
            if isinstance(p, list):
                runs = "".join(f'<w:r><w:rPr><w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:eastAsia="{f}"/><w:sz w:val="{sz}"/></w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>' for f, sz, t in p)
                return '<w:p><w:pPr><w:jc w:val="both"/></w:pPr>' + runs + '</w:p>'
            if arm == "head_th" and p.startswith("Article"):
                return '<w:p><w:r><w:rPr><w:b/><w:sz w:val="26"/></w:rPr><w:t xml:space="preserve">' + p + '</w:t></w:r></w:p>'
            return '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">' + p + '</w:t></w:r></w:p>'
        body = "".join(para(p) for p in paras)
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
               '<w:p><w:r><w:t>end</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/styles.xml", st)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
            if arm.endswith("_th"):
                z.writestr("word/_rels/document.xml.rels", drels.replace("</Relationships>", '<Relationship Id="rIdTh" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>'))
                z.writestr("word/theme/theme1.xml", theme)
            else:
                z.writestr("word/_rels/document.xml.rels", drels)
    print("ok", len(ARMS))


def lines_word(pdf):
    import fitz
    page = fitz.open(pdf)[0]; rows = {}
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                if s["text"].strip():
                    y = round(s["origin"][1], 2); rows.setdefault(y, {})
                    rows[y][s["font"]] = (round(s["bbox"][1], 2), round(s["bbox"][3], 2), s["text"][:8])
    return [(y, rows[y]) for y in sorted(rows)]


def lines_oxi(exe, docx):
    t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
    g = json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]; rows = {}
    for x in g:
        if x["char"].strip():
            y = round(x["top"], 2); rows.setdefault(y, {})
            rows[y].setdefault(x["font_family"], [round(x["baseline"], 2), ""])
            rows[y][x["font_family"]][1] = (rows[y][x["font_family"]][1] + x["char"])[:8]
    return [(y, rows[y]) for y in sorted(rows)]


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str((OUT / f"{arm}.pdf").resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {arm: lines_word(OUT / f"{arm}.pdf") for arm in ARMS}
    for arm, v in res.items():
        print(f"== {arm}")
        prev = None
        for y, fonts in v:
            print(f"   y={y:7.2f} {'' if prev is None else f'pitch {y - prev:6.2f}':13s} {fonts}"); prev = y
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    for arm in ARMS:
        print(f"== {arm}")
        prev = None
        for y, fonts in lines_oxi(os.path.abspath(exe), OUT / f"{arm}.docx"):
            print(f"   top={y:7.2f} {'' if prev is None else f'pitch {y - prev:6.2f}':13s} {fonts}"); prev = y


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
