# -*- coding: utf-8 -*-
"""Which face draws text set in the MISSING Arial Unicode MS, on this machine?

en/policies__0097fbf2: a numbering level whose rPr is rFonts ascii/hAnsi/eastAsia
= Arial Unicode MS (installed on neither PC) with w:hint="eastAsia"; the list
numbers «1», «2», «3» were drawn in Calibri by the old PC's Word and in ＭＳ 明朝
by this PC's Word (the line 3.84 higher, 4 pages -> 3).  The memory law
aum_substitute_by_ja_lang_run (old PC) reads "one substitute per document: a
ja-lang AUM run makes AUM Latin half-width Japanese, otherwise Arial".

Arms <where>_<hint>_<theme>_<lang>: where = marker (numbering level rPr) / run;
hint = none / ea; theme = jpan (theme minor Calibri with a Jpan ＭＳ 明朝 slot and
rPrDefault eastAsiaTheme=minorEastAsia, the document's shape) / plain (rPrDefault
eastAsia Calibri, no theme); lang = en (w:lang eastAsia en-US) / ja (ja-JP).
Text «1 Ab 2» after a Calibri «Lead» run.  Read: the embedded program per character
(Word PDF; Oxi glyphs).

    python tools/metrics/_pb_aum_missing_face_gen.py gen
    python tools/metrics/_pb_aum_missing_face_gen.py word
    python tools/metrics/_pb_aum_missing_face_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_aum_missing_face"
AUM = "Arial Unicode MS"
ARMS = {f"{w}_{h}_{t}_{l}": (w, h, t, l, "none", None) for w in ("marker", "run") for h in ("none", "ea")
        for t in ("jpan", "plain") for l in ("en", "ja")}
# fontTable declarations (en/policies__0097fbf2 declares AUM charset 80 + altName Yu Gothic)
for ft in ("ja80", "ja80alt"):
    for t in ("jpan", "plain"):
        for h in ("none", "ea"):
            ARMS[f"marker_{h}_{t}_en_{ft}"] = ("marker", h, t, "en", ft, None)
# a second AUM run elsewhere drawing an EAST ASIAN character (☐ with / without hint, 漢):
# does it switch the document's one AUM substitute? (0097fbf2's AUM levels carry ☐ + hint)
for extra in ("box_ea", "box_none", "kanji_none"):
    ARMS[f"run_none_plain_en_x{extra}"] = ("run", "none", "plain", "en", "none", extra)
    ARMS[f"marker_none_plain_en_x{extra}"] = ("marker", "none", "plain", "en", "none", extra)
THEME = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">'
         '<a:themeElements><a:clrScheme name="c"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
         '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
         '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>'
         '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>'
         '<a:fontScheme name="f"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ ゴシック"/></a:majorFont>'
         '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ 明朝"/></a:minorFont></a:fontScheme>'
         '<a:fmtScheme name="m"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
         '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
         '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
         '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
         '</a:themeElements></a:theme>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (where, h, theme, lang, ft, extra) in ARMS.items():
        hint = ' w:hint="eastAsia"' if h == "ea" else ""
        ealang = "ja-JP" if lang == "ja" else "en-US"
        aum = f'<w:rFonts w:ascii="{AUM}" w:hAnsi="{AUM}" w:eastAsia="{AUM}"{hint}/>'
        if theme == "jpan":
            dflt = '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
        else:
            dflt = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>' + dflt +
                  f'<w:sz w:val="20"/><w:lang w:val="en-US" w:eastAsia="{ealang}" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + T.NS + '><w:abstractNum w:abstractNumId="0">'
                     f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/>'
                     f'<w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr><w:rPr>{aum}</w:rPr></w:lvl></w:abstractNum>'
                     '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
        if where == "marker":
            paras = "".join('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Lead line text</w:t></w:r></w:p>' for _ in range(3))
        else:
            paras = f'<w:p><w:r><w:t xml:space="preserve">Lead </w:t></w:r><w:r><w:rPr>{aum}</w:rPr><w:t xml:space="preserve">1 Ab 2</w:t></w:r></w:p>'
        if extra:
            xh = ' w:hint="eastAsia"' if extra == "box_ea" else ""
            xc = "漢" if extra == "kanji_none" else "☐"
            paras += (f'<w:p><w:r><w:t xml:space="preserve">Tail </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="{AUM}" '
                      f'w:hAnsi="{AUM}" w:eastAsia="{AUM}"{xh}/></w:rPr><w:t>{xc}</w:t></w:r></w:p>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + paras +
               '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
                          + ('<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' if theme == "jpan" else "") + '</Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
                                + ('<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' if theme == "jpan" else "") + '</Relationships>')
        if ft != "none":
            alt = '<w:altName w:val="Yu Gothic"/>' if ft == "ja80alt" else ""
            fonts = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts ' + T.NS + '>'
                     f'<w:font w:name="{AUM}">{alt}<w:panose1 w:val="020B0604020202020204"/><w:charset w:val="80"/><w:family w:val="swiss"/>'
                     '<w:pitch w:val="variable"/><w:sig w:usb0="F7FFAFFF" w:usb1="E9DFFFFF" w:usb2="0000003F" w:usb3="00000000" w:csb0="003F01FF" w:csb1="00000000"/></w:font></w:fonts>')
            ct = ct.replace("</Types>", '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>')
            drels = drels.replace("</Relationships>", '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>')
        else:
            fonts = None
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            if fonts:
                z.writestr("word/fontTable.xml", fonts)
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml", numbering); z.writestr("word/document.xml", doc)
            if theme == "jpan":
                z.writestr("word/theme/theme1.xml", THEME)
    print("ok", len(ARMS))


def short(f):
    return (f or "?").lstrip("=")[:14]


def faces_word(pdf):
    import face_census as F
    got = {}
    for pi, c, x, y, face in F.word_chars(pdf, 1):
        if c in "12Ab":
            got.setdefault(c, short(face))
    return got


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
    res = {arm: faces_word(str(OUT / f"{arm}.pdf")) for arm in ARMS}
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")
    for arm, g in res.items():
        print(f"{arm:22s}", g)


def oxi(exe):
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        got = {}
        for g in json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]:
            if g["char"] in "12Ab":
                got.setdefault(g["char"], short(g.get("font_family")))
        print(f"{arm:22s} word {wd[arm]}  oxi {got}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
