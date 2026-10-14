# -*- coding: utf-8 -*-
"""Where does a compat-12 (Word 2007 mode) document put a one-line footer?

blind-G ja/reports__597cb86c and 59eadfc4 (the same FullPicture.app template):
footer = one Calibri 12 paragraph with <w:pBdr><w:top sz=4/> (no space attr),
no pStyle, Normal style empty, no pPrDefault, footer 720 / bottom 1440, page
height 16837.795 twips, evenAndOddHeaders on, compatibilityMode 12.  Word's
baseline is 792.34 (13.55 above the footer line at 805.89); Oxi's is 793.35.

Arms c<compat>_<bdr>_<eo>: compat 12 / 14 / 15; bdr = none / top4 (sz 4, no
space); eo = eo (evenAndOddHeaders) / no.  Everything else mirrors the
document.  Read: the footer baseline (Word PDF; Oxi glyph dump) and the first
body baseline.

    python tools/metrics/_pb_ftr_compat12_gen.py gen
    python tools/metrics/_pb_ftr_compat12_gen.py word
    python tools/metrics/_pb_ftr_compat12_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_ftr_compat12"
BDRS = {"none": "", "top4": '<w:pBdr><w:top w:val="single" w:sz="4"/></w:pBdr>'}
ARMS = [f"c{c}_{b}_{e}" for c in (12, 14, 15) for b in BDRS for e in ("eo", "no")]
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')


def footer(bdr):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + T.NS + '><w:p><w:pPr>' + BDRS[bdr] +
            '</w:pPr><w:r><w:rPr/><w:t xml:space="preserve">Footer created by </w:t></w:r><w:r><w:rPr><w:color w:val="2980b9"/>'
            '<w:u w:val="single"/></w:rPr><w:t>FullPicture.app</w:t></w:r></w:p></w:ftr>')


def settings(compat, eo):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '>' +
            ('<w:evenAndOddHeaders/>' if eo == "eo" else '') +
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="%d"/></w:compat></w:settings>' % compat)


def doc_xml():
    body = "".join('<w:p><w:r><w:t>Body line %d of the filler text that keeps the page full.</w:t></w:r></w:p>' % (i + 1) for i in range(30))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + body +
            '<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:orient="portrait" w:w="11905.511811023622" w:h="16837.79527559055"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:num="1" w:space="720"/></w:sectPr></w:body></w:document>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm in ARMS:
        c, b, e = arm.split("_")
        ct = T.CT.replace("</Types>", '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
                                '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings(int(c[1:]), e))
            z.writestr("word/footer1.xml", footer(b)); z.writestr("word/document.xml", doc_xml())
    print("ok", len(ARMS))


def _read(lines):
    f = [l["y"] for l in lines if l["t"].startswith("Footer")]
    body = sorted(l["y"] for l in lines if l["t"].startswith("Body"))
    return {"ftr": round(f[0], 2) if f else None, "body_first": round(body[0], 2) if body else None, "n_body": len(body)}


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
        print(f"{arm:14s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        wd = word.get(arm) or {}
        flag = "" if o["ftr"] is not None and wd.get("ftr") is not None and abs(o["ftr"] - wd["ftr"]) < 0.6 else "  <-- differs"
        print(f"{arm:14s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
