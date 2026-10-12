# -*- coding: utf-8 -*-
"""Does a footer paragraph's space-after survive when the NEXT footer paragraph
has an EXACT line rule?

blind-H en/technical__0186a34e (first-page footer): «Footer» para (Calibri 7,
after=80) then three «Footer 2» paras (Calibri 6, line 140 exact, after 0).
Word: line 1 → line 2 = 11.40 (= descent 1.75 + after 4 + 0.8 x 7); Oxi 7.48
(the 4pt gone), so line 1 sits 3.8 low while lines 2-4 match.

Arms: <where>_<b>: where = ftr / body; b = exact (line 140 exact), auto
(single), atleast (line 140 atLeast), exact0 (exact but A has after=0).
Read baselines of «Alpha» (A) and «Beta» (B) and the gap.

    python tools/metrics/_pb_ftrafter_exact_gen.py gen
    python tools/metrics/_pb_ftrafter_exact_gen.py word
    python tools/metrics/_pb_ftrafter_exact_gen.py oxi <renderer.exe>
"""
import json, os, shutil, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_ftrafter_exact"
BS = {"exact": '<w:spacing w:after="0" w:line="140" w:lineRule="exact"/>', "auto": '<w:spacing w:after="0"/>',
      "atleast": '<w:spacing w:after="0" w:line="140" w:lineRule="atLeast"/>', "exact0": '<w:spacing w:after="0" w:line="140" w:lineRule="exact"/>',
      # lastafter: single lines, and the LAST paragraph carries after=80 -- does a trailing after lift the stack?
      "lastafter": '<w:spacing w:after="0"/>'}
WHERES = ("ftr", "body")
ARMS = [f"{w}_{b}" for w in WHERES for b in BS]
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:widowControl w:val="0"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:tabs><w:tab w:val="right" w:pos="10065"/></w:tabs><w:spacing w:after="80"/></w:pPr><w:rPr><w:sz w:val="14"/></w:rPr></w:style>'
          '</w:styles>')


def paras(b):
    a_sp = '<w:spacing w:after="0"/>' if b == "exact0" else ""
    c_sp = '<w:spacing w:after="80"/>' if b == "lastafter" else BS[b]
    return (f'<w:p><w:pPr><w:pStyle w:val="Footer"/>{a_sp}</w:pPr><w:r><w:t>Alpha first line of the stack</w:t></w:r></w:p>'
            f'<w:p><w:pPr><w:pStyle w:val="Footer"/>{BS[b]}<w:rPr><w:sz w:val="12"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="12"/></w:rPr><w:t>Beta second line</w:t></w:r></w:p>'
            f'<w:p><w:pPr><w:pStyle w:val="Footer"/>{c_sp}<w:rPr><w:sz w:val="12"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="12"/></w:rPr><w:t>Gamma third line</w:t></w:r></w:p>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    for arm in ARMS:
        where, b = arm.split("_")
        body = "".join(f'<w:p><w:r><w:t>Body line {i + 1}.</w:t></w:r></w:p>' for i in range(5))
        ct, drels, parts = T.CT, T.DRELS, {}
        if where == "ftr":
            ct = ct.replace("</Types>", '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
            drels = drels.replace("</Relationships>", '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
            parts["word/footer1.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + T.NS + '>' + paras(b) + '</w:ftr>'
            ref = '<w:footerReference w:type="default" r:id="rIdF"/>'
        else:
            body += paras(b); ref = ""
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + "><w:body>" + body +
               '<w:sectPr>' + ref + '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="2268"'
               ' w:left="1134" w:header="731" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            for n, c in parts.items():
                z.writestr(n, c)
            z.writestr("word/document.xml", doc)
    print("ok", len(ARMS))


def _read(lines):
    g = lambda k: next((l["y"] for l in lines if l["t"].startswith(k)), None)
    a, b, c = g("Alpha"), g("Beta"), g("Gamma")
    return {"A": round(a, 2) if a else None, "B": round(b, 2) if b else None, "C": round(c, 2) if c else None,
            "AB": round(b - a, 2) if a and b else None, "BC": round(c - b, 2) if b and c else None}


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
        print(f"{arm:13s} word {res[arm]}", flush=True)
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")


def oxi(exe):
    import line_census as L
    word = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    for arm in ARMS:
        o = _read(L.oxi_lines(os.path.abspath(exe), OUT / f"{arm}.docx"))
        wd = word.get(arm) or {}
        flag = "" if o["AB"] is not None and wd.get("AB") is not None and abs(o["AB"] - wd["AB"]) < 0.4 else "  <-- differs"
        print(f"{arm:13s} oxi {o}  word {wd}{flag}", flush=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
