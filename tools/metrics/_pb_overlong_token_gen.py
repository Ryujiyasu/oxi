# -*- coding: utf-8 -*-
"""An over-long token (wider than a full line) that starts MID-line: does Word
pack it on the current line to the last fitting character (S1059, derived from
Latin URLs), or move it whole to the next line first?  blind-G ja/reports__
59eadfc4 (compat 12, docDefaults eastAsia Calibri → Japanese/Chinese wraps
like Latin, d4fc04ad): Word leaves «2. » alone on its line and starts the
unbreakable Chinese run on the next one, then breaks it at the margin; Oxi
packs it after «2. ».  Arms (12pt, A4, margins 1440, Normal.dotm defaults):
    cjk_c12       eastAsia Calibri, compat 12, «2. » + 70 CJK (no spaces)
    cjk_c15       the same at compat 15 (character wrap → packs)
    cjk_ms_c12    eastAsia ＭＳ 明朝, compat 12 (character wrap → packs)
    lat_c12       eastAsia Calibri, compat 12, «2. » + 90 Latin letters
    lat_c15       the same at compat 15
    cjk_sp_c12    «2. » + 20 CJK + space + 70 CJK (the long token after a real word)
    cjk_dig_c12 / cjk_com_c12 / cjk_dig_c15 / lat_dig_c12 / cjk_digsp_c12 / cjk_kana_c12
                  digits «2000», a «，», or « 2000 » inside the token (the document's «获得2000积分…，但»)
Read: the text of lines 1-3 (Word PDF via fitz; Oxi via --dump-glyphs).

    python tools/metrics/_pb_overlong_token_gen.py gen
    python tools/metrics/_pb_overlong_token_gen.py word
    python tools/metrics/_pb_overlong_token_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402
import line_census as L  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_overlong_token"
CJK = "会员申请通过后可以获得积分的限时奖励但必须从指定页面申请才能享受此优惠会员申请通过后可以获得积分的限时奖励但必须从指定页面申请才能享受此优惠"
LAT = "abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzabcdefghijkl"
ARMS = {"cjk_c12": ("Calibri", 12, "2. " + CJK), "cjk_c15": ("Calibri", 15, "2. " + CJK),
        "cjk_ms_c12": ("ＭＳ 明朝", 12, "2. " + CJK), "lat_c12": ("Calibri", 12, "2. " + LAT),
        "lat_c15": ("Calibri", 15, "2. " + LAT), "cjk_sp_c12": ("Calibri", 12, "2. " + CJK[:20] + " " + CJK),
        # the document's token has digits and a comma inside: are those break opportunities under legacy wrap?
        "cjk_dig_c12": ("Calibri", 12, "2. " + CJK[:10] + "2000" + CJK), "cjk_com_c12": ("Calibri", 12, "2. " + CJK[:10] + "，" + CJK),
        "cjk_dig_c15": ("Calibri", 15, "2. " + CJK[:10] + "2000" + CJK), "lat_dig_c12": ("Calibri", 12, "2. " + LAT[:10] + "2000" + LAT),
        "cjk_digsp_c12": ("Calibri", 12, "2. " + CJK[:10] + " 2000 " + CJK), "cjk_kana_c12": ("Calibri", 12, "2. " + "あいうえおかきくけこ" + "2000" + CJK),
        # the document's paragraph verbatim: jc=both, and with its theme (ＭＳ 明朝 fallback in Oxi)
        "doc_jc": ("Calibri", 12, None), "doc_jc_th": ("Calibri", 12, None), "doc_nojc": ("Calibri", 12, None),
        # the theme alone, with the plain / digit / comma tokens
        "cjk_c12_th": ("Calibri", 12, "2. " + CJK), "cjk_dig_c12_th": ("Calibri", 12, "2. " + CJK[:10] + "2000" + CJK),
        "cjk_com_c12_th": ("Calibri", 12, "2. " + CJK[:10] + "，" + CJK), "cjk_sp_c12_th": ("Calibri", 12, "2. " + CJK[:20] + " " + CJK)}
DOC = "2. 会员申请通过后可以获得2000积分的限时奖励，但必须从指定页面申请才能享受此优惠。"
ARMS = {k: (v[0], v[1], v[2] if v[2] is not None else DOC) for k, v in ARMS.items()}
THEME_SRC = REPO / "pipeline_data" / "docx_corpus" / "ja" / "reports" / "59eadfc480bf174c.docx"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (ea, compat, text) in ARMS.items():
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  f'<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="{ea}" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/>'
                  '</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body><w:p><w:r><w:t xml:space="preserve">' + text + '</w:t></w:r></w:p>'
               '<w:p><w:r><w:t>end</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="' + str(compat) + '"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                          '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        if arm.endswith("_th"):
            drels = drels.replace("</Relationships>", '<Relationship Id="rIdTh" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
            if arm.endswith("_th"):
                z.writestr("word/theme/theme1.xml", zipfile.ZipFile(THEME_SRC).read("word/theme/theme1.xml"))
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
    return [(round(l["y"], 1), len(l["t"]), l["t"][:6] + "…" + l["t"][-4:]) for l in L.group_lines(items)[:3]]


def lines_oxi(exe, docx):
    t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
    g = json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]
    items = [(x["baseline"], x["x"], x["x"] + x["font_size"], x["char"]) for x in g if x["char"].strip()]
    return [(round(l["y"], 1), len(l["t"]), l["t"][:6] + "…" + l["t"][-4:]) for l in L.group_lines(items)[:3]]


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
        print(f"{arm:12s} word {v}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = lines_oxi(os.path.abspath(exe), OUT / f"{arm}.docx")
        flag = "" if [x[1] for x in o] == [x[1] for x in wd[arm]] else "  <-- differs"
        print(f"{arm:12s} oxi  {o}\n{'':12s} word {wd[arm]}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
