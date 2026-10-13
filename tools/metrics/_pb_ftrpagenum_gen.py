# -*- coding: utf-8 -*-
"""Where does Word put a footer's framePr page number (the classic
「ページ番号」 frame: framePr wrap=around vAnchor=text hAnchor=margin
xAlign=center y=1) when the footer also holds an empty paragraph after it?

blind-H ja/legal__0f631d46 (page 792 high, footer dist 36, bottom 85.05; footer
= frame paragraph «PAGE» + an empty paragraph, footer style snapToGrid=0,
docDefaults 12pt theme 游明朝): Word's «1» baseline 753.60, Oxi 751.38.

Arms (one docx each, one body paragraph; read the page number's baseline
measured UP from the page bottom): faithful (frame + empty para, snapToGrid 0),
noempty (frame paragraph only), snapon (no snapToGrid 0), y0 (frame y=0),
msmin (ＭＳ 明朝 10.5 instead of the theme face), twoempty (two empty paras).

    python tools/metrics/_pb_ftrpagenum_gen.py gen
    python tools/metrics/_pb_ftrpagenum_gen.py word
    python tools/metrics/_pb_ftrpagenum_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_ftrpagenum"
PAGE_H = 792.0
ARMS = ["faithful", "noempty", "snapon", "y0", "msmin", "twoempty", "century", "century_noempty"]
THEME = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">'
         '<a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
         '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
         '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
         '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
         '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="游ゴシック Light"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/></a:majorFont>'
         '<a:minorFont><a:latin typeface="游明朝"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游明朝"/></a:minorFont></a:fontScheme>'
         '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
         '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
         '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
         '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
         '</a:themeElements></a:theme>')


def styles(snap0, msmin):
    rfonts = ('<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/>' if msmin
              else '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="24"/><w:szCs w:val="24"/>')
    snap = '<w:snapToGrid w:val="0"/>' if snap0 else ""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>' + rfonts +
            '<w:kern w:val="2"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="a8"><w:name w:val="footer"/><w:basedOn w:val="a"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4252"/><w:tab w:val="right" w:pos="8504"/></w:tabs>' + snap + '</w:pPr></w:style>'
            '<w:style w:type="character" w:styleId="ad"><w:name w:val="page number"/><w:basedOn w:val="a0"/></w:style>'
            '<w:style w:type="character" w:default="1" w:styleId="a0"><w:name w:val="Default Paragraph Font"/></w:style></w:styles>')


def footer(empties, y):
    frame = (f'<w:p><w:pPr><w:pStyle w:val="a8"/><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="center" w:y="{y}"/>'
             '<w:rPr><w:rStyle w:val="ad"/></w:rPr></w:pPr><w:r><w:rPr><w:rStyle w:val="ad"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
             '<w:r><w:rPr><w:rStyle w:val="ad"/></w:rPr><w:instrText xml:space="preserve">PAGE  </w:instrText></w:r>'
             '<w:r><w:rPr><w:rStyle w:val="ad"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rStyle w:val="ad"/><w:noProof/></w:rPr><w:t>1</w:t></w:r>'
             '<w:r><w:rPr><w:rStyle w:val="ad"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + T.NS + '>' + frame
            + '<w:p><w:pPr><w:pStyle w:val="a8"/></w:pPr></w:p>' * empties + '</w:ftr>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm in ARMS:
        theme = THEME.replace('typeface="游明朝"', 'typeface="Century"', 1).replace('<a:font script="Jpan" typeface="游明朝"/></a:minorFont>', '<a:font script="Jpan" typeface="ＭＳ 明朝"/></a:minorFont>') if arm.startswith('century') else THEME
        empties = {"noempty": 0, "twoempty": 2, "century_noempty": 0}.get(arm, 1)
        y = 0 if arm == "y0" else 1
        body = '<w:p><w:r><w:t>本文の段落です。</w:t></w:r></w:p>'
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701"'
               ' w:left="1701" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>')
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                          '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                                '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
                                '<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>')
        settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:defaultTabStop w:val="840"/>'
                    '<w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:useFELayout/>'
                    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", styles(arm != "snapon", arm == "msmin"))
            z.writestr("word/settings.xml", settings); z.writestr("word/theme/theme1.xml", theme)
            z.writestr("word/footer1.xml", footer(empties, y)); z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def _read(lines):
    ls = [l for l in lines if l["t"].strip() == "1"]
    return {"from_bottom": round(PAGE_H - ls[0]["y"], 2), "x": round(ls[0]["x0"], 1)} if ls else None


def word():
    import win32com.client, line_census as L
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
        res[arm] = _read(L.word_lines(OUT / f"{arm}.pdf"))
        print(f"{arm:9s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o, wd = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx")), word.get(arm)
        flag = "" if o and wd and abs(o["from_bottom"] - wd["from_bottom"]) < 0.5 else "  <-- differs"
        print(f"{arm:9s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
