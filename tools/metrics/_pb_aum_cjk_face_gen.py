# -*- coding: utf-8 -*-
"""Which face draws CJK text whose run font is missing on this machine, when the
fontTable names Arial Unicode MS (also missing here) as the altName?  blind-G
ja/correspondence__139e74eb: the title run is «点字線なし» (a Braille face,
fontTable charset 80, altName Arial Unicode MS); Word's PDF draws its kana in
YuGothic-Regular (baseline 117.02 = line top + 14.18), Oxi applies the altName
and sizes the line by the AUM/ＭＳ 明朝 box (114.56).  Arms (12pt, A4):
    tenji_alt     «点字線なし» with fontTable altName AUM (the document)
    tenji_noalt   «点字線なし» with a fontTable entry but no altName
    tenji_noft    «点字線なし» without any fontTable entry
    aum_direct    the run declares Arial Unicode MS, fontTable entry charset 80
    aum_noft      the run declares Arial Unicode MS, no fontTable
    ms_ctrl       ＭＳ 明朝 (control)
Each arm: two paragraphs «濁しりつかつ濁とーつーしん» then «end» in Calibri.
Read: the PDF span fonts and baselines (Word); the dump faces and baselines (Oxi).

    python tools/metrics/_pb_aum_cjk_face_gen.py gen
    python tools/metrics/_pb_aum_cjk_face_gen.py word
    python tools/metrics/_pb_aum_cjk_face_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_aum_cjk_face"
TEXT = "濁しりつかつ濁とーつーしん"
FT_TENJI = '<w:font w:name="点字線なし">{alt}<w:panose1 w:val="02000604000000000000"/><w:charset w:val="80"/><w:family w:val="auto"/><w:pitch w:val="variable"/><w:sig w:usb0="A00002BF" w:usb1="68C7FCFB" w:usb2="00000010" w:usb3="00000000" w:csb0="0002009F" w:csb1="00000000"/></w:font>'
FT_AUM = '<w:font w:name="Arial Unicode MS"><w:panose1 w:val="020B0604020202020204"/><w:charset w:val="80"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="FFFFFFFF" w:usb1="E9FFFFFF" w:usb2="0000003F" w:usb3="00000000" w:csb0="003F01FF" w:csb1="00000000"/></w:font>'
# (face, fontTable xml or None)
ARMS = {"tenji_alt": ("点字線なし", FT_TENJI.format(alt='<w:altName w:val="Arial Unicode MS"/>')),
        "tenji_noalt": ("点字線なし", FT_TENJI.format(alt="")), "tenji_noft": ("点字線なし", None),
        "aum_direct": ("Arial Unicode MS", FT_AUM), "aum_noft": ("Arial Unicode MS", None), "ms_ctrl": ("ＭＳ 明朝", None)}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm, (face, ft) in ARMS.items():
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="ＭＳ 明朝" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/>'
                  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>'
               f'<w:p><w:r><w:rPr><w:rFonts w:ascii="{face}" w:eastAsia="{face}" w:hAnsi="{face}" w:hint="eastAsia"/></w:rPr><w:t>{TEXT}</w:t></w:r></w:p>'
               '<w:p><w:r><w:t>end</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                          + ('<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>' if ft else '') + '</Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                                + ('<Relationship Id="rIdFt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>' if ft else '') + '</Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles)
            z.writestr("word/settings.xml", settings); z.writestr("word/document.xml", doc)
            if ft:
                z.writestr("word/fontTable.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts ' + T.NS + '>' + ft + '</w:fonts>')
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
        pg = fitz.open(str(OUT / f"{arm}.pdf"))[0]; rows = []
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip():
                        rows.append((round(s["origin"][1], 2), s["font"], round(s["ascender"], 3), s["text"][:8]))
        res[arm] = rows; print(f"{arm:12s} word {rows}")
    (OUT / "word.json").write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    for arm in ARMS:
        t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True, timeout=600)
        g = json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]; rows = {}
        for x in g:
            if x["char"].strip():
                k = (round(x["baseline"], 2), x["font_family"]); rows.setdefault(k, ""); rows[k] = (rows[k] + x["char"])[:8]
        print(f"{arm:12s} oxi  {[(k[0], k[1], v) for k, v in sorted(rows.items())]}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
