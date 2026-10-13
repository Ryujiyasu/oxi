# -*- coding: utf-8 -*-
"""When does Word wrap Japanese text between characters, and when does it wrap
it like Latin, only at spaces?  blind-G ja/reports__59eadfc4 (docDefaults
Calibri with eastAsia="Calibri", lang val="en-US" and no eastAsia language,
settings themeFontLang en-US): the space-separated title «【楽天市場】超目玉！半額！
抱っこ 防寒 ベビーケープ 防寒ケープ ダウン フットカバー付き …» wraps in Word before
«フットカバー付き» (409pt used of 451) while Oxi breaks inside the word and
fits three more characters. Arms (12pt, A4, margins 1440, the same text):
    cal_nolang   eastAsia Calibri, lang en-US only                (the document)
    cal_ja       eastAsia Calibri, lang eastAsia ja-JP
    ms_nolang    eastAsia ＭＳ 明朝, lang en-US only
    ms_ja        eastAsia ＭＳ 明朝, lang eastAsia ja-JP
    cal_theme_ja eastAsia Calibri, lang en-US, settings themeFontLang eastAsia ja-JP
    cal_nospace  eastAsia Calibri, lang en-US, the text without spaces
Read: the text of line 1 and line 2 (Word PDF via fitz; Oxi via --dump-glyphs).

    python tools/metrics/_pb_cjkwrap_lang_gen.py gen
    python tools/metrics/_pb_cjkwrap_lang_gen.py word
    python tools/metrics/_pb_cjkwrap_lang_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402
import line_census as L  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_cjkwrap_lang"
TEXT = "【楽天市場】超目玉！半額！抱っこ 防寒 ベビーケープ 防寒ケープ ダウン フットカバー付き あったか マルチプル マルチケープ フリース 手書きドット柄 くま耳付き はっ水 秋冬 暖かい 出産祝い ギフト エルゴ対応：kidsmio マット・ベビーサークル"
# (eastAsia face, run lang eastAsia attr, themeFontLang eastAsia attr, strip spaces, compatibilityMode)
ARMS = {"cal_nolang": ("Calibri", "", "", False, 15), "cal_ja": ("Calibri", ' w:eastAsia="ja-JP"', "", False, 15),
        "ms_nolang": ("ＭＳ 明朝", "", "", False, 15), "ms_ja": ("ＭＳ 明朝", ' w:eastAsia="ja-JP"', "", False, 15),
        "cal_theme_ja": ("Calibri", "", ' w:eastAsia="ja-JP"', False, 15), "cal_nospace": ("Calibri", "", "", True, 15),
        # 59eadfc4 is compatibilityMode 12 (characterSpacingControl doNotCompress): the legacy modes
        "cal_nolang_c12": ("Calibri", "", "", False, 12), "cal_ja_c12": ("Calibri", ' w:eastAsia="ja-JP"', "", False, 12),
        "ms_nolang_c12": ("ＭＳ 明朝", "", "", False, 12), "ms_ja_c12": ("ＭＳ 明朝", ' w:eastAsia="ja-JP"', "", False, 12),
        "cal_nolang_c14": ("Calibri", "", "", False, 14), "cal_nospace_c12": ("Calibri", "", "", True, 12)}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (ea, lang_ea, theme_ea, nospace, compat) in ARMS.items():
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  f'<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="{ea}" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"{lang_ea}/>'
                  '</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        text = TEXT.replace(" ", "") if nospace else TEXT
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body><w:p><w:r><w:t xml:space="preserve">' + text + '</w:t></w:r></w:p>'
               '<w:p><w:r><w:t>end</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="' + str(compat) + '"/></w:compat>'
                    f'<w:themeFontLang w:val="en-US"{theme_ea}/></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def lines_word(pdf):
    import fitz
    page = fitz.open(pdf)[0]; items = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"].strip():
                        items.append((c["origin"][1], c["origin"][0], c["bbox"][2], c["c"]))
    return [(round(l["y"], 1), len(l["t"]), l["t"][-8:]) for l in L.group_lines(items)[:3]]


def lines_oxi(exe, docx):
    t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
    g = json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]
    items = [(x.get("baseline") or x["top"] + L.win_ascent(x.get("font_family", "")) * x["font_size"], x["x"], x["x"] + x["font_size"], x["char"]) for x in g if x["char"].strip()]
    return [(round(l["y"], 1), len(l["t"]), l["t"][-8:]) for l in L.group_lines(items)[:3]]


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
        print(f"{arm:13s} word {v}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = lines_oxi(os.path.abspath(exe), OUT / f"{arm}.docx")
        flag = "" if [x[1] for x in o] == [x[1] for x in wd[arm]] else "  <-- differs"
        print(f"{arm:13s} oxi  {o}\n{'':13s} word {wd[arm]}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
