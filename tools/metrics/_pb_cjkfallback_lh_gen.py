# -*- coding: utf-8 -*-
"""(pPrDefault after=0 line=240 is explicit: without it this machine's Normal.dotm adds after=160 / line=278.)

Line height of a CJK line whose eastAsia face is Latin-only (Calibri, from
docDefaults): which face sets the LINE while glyphs fall back per character?

blind-G ja/reports__597cb86c / 59eadfc4 (docDefaults eastAsia=Calibri, theme
minor Jpan=ＭＳ 明朝, themeFontLang en-US, Simplified-Chinese body): Word's PDF
draws the glyphs in MS-Mincho (JIS chars) / SimSun (Simplified-only chars) /
YuGothic («类»), all on one baseline, and every such line is 20.04 tall = 12pt
× 1.287 (the Yu usWin box) × 83/64 — not ＭＳ 明朝's 15.56. Oxi draws the glyphs
in the same faces but sizes the line from them (15.56 / 23.08).

One docx per arm, two identical 12pt paragraphs each (the pitch between their
baselines is the line height), docDefaults eastAsia=Calibri, lang en-US:
    ja_mincho   theme Jpan=ＭＳ 明朝, text 当然的事情 (JIS chars)
    hans_mincho theme Jpan=ＭＳ 明朝, text 这访对选择 (Simplified-only chars)
    mix_mincho  theme Jpan=ＭＳ 明朝, text 1. 当访问隐私 cookie 存储 (the doc's shape)
    ja_yumin    theme Jpan=游明朝, text 当然的事情
    ja_notheme  no theme part, text 当然的事情
    mix_jajp    theme Jpan=ＭＳ 明朝 + themeFontLang ja-JP, mixed text
    latin_only  theme Jpan=ＭＳ 明朝, text Latin only (control: Calibri 14.65)
Read: baseline pitch, and the span fonts of the first line.

    python tools/metrics/_pb_cjkfallback_lh_gen.py gen
    python tools/metrics/_pb_cjkfallback_lh_gen.py word
    python tools/metrics/_pb_cjkfallback_lh_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_cjkfallback_lh"
JA, HANS, MIX, LAT = "当然的事情", "这访对选择", "1. 当访问隐私 cookie 存储和检索", "Article summary text"
# (theme Jpan face or None for no theme, text, themeFontLang)
ARMS = {
    "ja_mincho": ("ＭＳ 明朝", JA, "en-US"), "hans_mincho": ("ＭＳ 明朝", HANS, "en-US"), "mix_mincho": ("ＭＳ 明朝", MIX, "en-US"),
    "ja_yumin": ("游明朝", JA, "en-US"), "ja_notheme": (None, JA, "en-US"), "mix_jajp": ("ＭＳ 明朝", MIX, "ja-JP"),
    "latin_only": ("ＭＳ 明朝", LAT, "en-US"),
}
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/>'
          '<w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
THEME = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">'
         '<a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
         '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
         '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
         '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
         '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="{JPAN}"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/></a:majorFont>'
         '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="{JPAN}"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/></a:minorFont></a:fontScheme>'
         '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
         '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
         '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
         '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
         '</a:themeElements></a:theme>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (jpan, text, lang) in ARMS.items():
        body = "".join(f'<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:rPr/><w:t xml:space="preserve">Ｑ{i}{text}</w:t></w:r></w:p>' for i in range(2))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
               ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        has_theme = jpan is not None
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                          + ('<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' if has_theme else "") + "</Types>")
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                                + ('<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' if has_theme else "") + "</Relationships>")
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="720"/>'
                    f'<w:themeFontLang w:val="{lang}"/><w:compat>'
                    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
            if has_theme:
                z.writestr("word/theme/theme1.xml", THEME.replace("{JPAN}", jpan))
    print("ok", len(ARMS))


def _read(lines):
    ls = sorted([l for l in lines if "Ｑ" in l["t"]], key=lambda l: l["y"])
    return {"pitch": round(ls[1]["y"] - ls[0]["y"], 2) if len(ls) >= 2 else None, "y0": round(ls[0]["y"], 2) if ls else None}


def word():
    import win32com.client, line_census as L, fitz
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
        r = _read(L.word_lines(OUT / f"{arm}.pdf"))
        pg = fitz.open(OUT / f"{arm}.pdf")[0]
        fonts = []
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip() and r["y0"] and abs(s["origin"][1] - r["y0"]) < 1.0:
                        fonts.append(f'{s["font"]}:{s["text"].strip()[:4]}')
        r["fonts"] = fonts[:8]
        res[arm] = r
        print(f"{arm:12s} word {r}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        wd = word.get(arm, {})
        flag = "" if o["pitch"] and wd.get("pitch") and abs(o["pitch"] - wd["pitch"]) < 0.3 else "  <-- differs"
        print(f"{arm:12s} oxi {o}  word pitch {wd.get('pitch')} fonts {wd.get('fonts')}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
