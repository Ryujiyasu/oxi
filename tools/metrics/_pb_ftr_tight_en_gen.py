# -*- coding: utf-8 -*-
"""How far may a LATIN body line reach into a footer-constrained page bottom?

S726 (2026-07-03) gives a page whose footer stands above the bottom margin NO
leniency: the line's FULL box must fit above the footer top.  S1542 measured
the CJK body side (Word lets the box overhang the footer top by the footer's
first paragraph win descent) and left the Latin side to "an EN probe" that was
never written.  blind-G corpus en/policies__008f68d9 (Legal, margins 36, footer
36 + Cambria 11 text + empty paragraph = 25.8; Bookman Old Style 10 body at
line 276): Word keeps a line whose 1.15x box ends 0.14 past the footer top,
Oxi pushes it -- the last paragraph of page 2 then lands on page 3.

Arms <font>_<mult>_<ftr>_b<before tenths>: font = ari (Arial 10) / bkm (Bookman
Old Style 10); mult = 240 / 276 auto; ftr = txt (Cambria 11 text + empty para,
the document's footer) / none; the first body paragraph's space-before sweeps
0..14pt in 0.5pt steps so the last page-1 line crosses the limit.
Read: the page-1 body line count and the last page-1 body baseline (Word PDF;
Oxi glyphs) -> the flip point per configuration.

    python tools/metrics/_pb_ftr_tight_en_gen.py gen
    python tools/metrics/_pb_ftr_tight_en_gen.py word
    python tools/metrics/_pb_ftr_tight_en_gen.py oxi <renderer.exe>
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _pb_txbxfit_gen as T  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_pb_ftr_tight_en"
FONTS = {"ari": "Arial", "bkm": "Bookman Old Style"}
BEFORES = [i * 5 for i in range(29)]  # tenths of a point: 0.0 .. 14.0
CONFIGS = [(f, m, ftr) for f in FONTS for m in (240, 276) for ftr in ("txt", "none")]
ARMS = {f"{f}_{m}_{ftr}_b{b:03d}": (f, m, ftr, b) for (f, m, ftr) in CONFIGS for b in BEFORES}
NLINES = 90


def footer_xml():
    cam = '<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:sz w:val="22"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + T.NS + '>'
            f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:rPr>{cam}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{cam}</w:rPr><w:t>This institution is an equal opportunity provider.</w:t></w:r></w:p>'
            f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>{cam}</w:rPr></w:pPr></w:p></w:ftr>')


def doc_xml(face, mult, ftr, before_tenths):
    rpr = f'<w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}"/><w:sz w:val="20"/></w:rPr>'
    sp = lambda b: f'<w:spacing w:before="{b}" w:after="0" w:line="{mult}" w:lineRule="auto"/>'
    first = f'<w:p><w:pPr>{sp(before_tenths * 2)}{rpr}</w:pPr><w:r>{rpr}<w:t>B00 first body line of the sweep</w:t></w:r></w:p>'
    rest = "".join(f'<w:p><w:pPr>{sp(0)}{rpr}</w:pPr><w:r>{rpr}<w:t>B{i:02d} body line of the sweep</w:t></w:r></w:p>' for i in range(1, NLINES))
    ref = '<w:footerReference w:type="default" r:id="rIdF"/>' if ftr == "txt" else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + T.NS + '><w:body>' + first + rest +
            '<w:sectPr>' + ref + '<w:pgSz w:w="12240" w:h="20160" w:code="5"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + T.NS + '><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:lang w:val="en-US"/>'
          '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + T.NS + '><w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    for arm, (f, m, ftr, b) in ARMS.items():
        ct = T.CT.replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
        drels = T.DRELS.replace("</Relationships>", '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
                                '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
        with zipfile.ZipFile(OUT / f"{arm}.docx", "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", T.RELS)
            z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/styles.xml", STYLES)
            z.writestr("word/settings.xml", settings); z.writestr("word/footer1.xml", footer_xml())
            z.writestr("word/document.xml", doc_xml(FONTS[f], m, ftr, b))
    print("ok", len(ARMS))


def summarize(res, label):
    for cfg in CONFIGS:
        pre = "%s_%d_%s_" % cfg
        rows = [(int(k.rsplit("_b", 1)[1]) / 10.0, v) for k, v in sorted(res.items()) if k.startswith(pre)]
        flips = [(rows[i][0], rows[i][1], rows[i + 1][1]) for i in range(len(rows) - 1) if rows[i][1]["n"] != rows[i + 1][1]["n"]]
        print(f"{label} {pre:16s} footer_base {rows[0][1].get('ftr')}  flips:", flush=True)
        for b, a, c in flips:
            print(f"      before {b:4.1f}: n {a['n']} last {a['last']}  -> before {b + 0.5:4.1f}: n {c['n']} last {c['last']}", flush=True)


def word():
    import win32com.client, fitz
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for arm in ARMS:
            pdf = OUT / f"{arm}.pdf"
            if pdf.exists():
                continue
            t = os.path.join(tempfile.mkdtemp(), arm + ".docx"); shutil.copy(OUT / f"{arm}.docx", t)
            d = w.Documents.Open(t, ReadOnly=False, AddToRecentFiles=False)
            d.SaveAs2(str(pdf.resolve()), 17); d.Close(0)
    finally:
        w.Quit()
    res = {}
    for arm in ARMS:
        pg = fitz.open(str(OUT / f"{arm}.pdf"))[0]; body = []; ftr = None
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"])
                y = round(l["spans"][0]["origin"][1], 2)
                if t.startswith("B"):
                    body.append(y)
                elif t.startswith("This institution"):
                    ftr = y
        res[arm] = {"n": len(body), "last": max(body) if body else None, "ftr": ftr}
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")
    summarize(res, "word")


def oxi(exe):
    res = {}
    for arm in ARMS:
        t = tempfile.mkdtemp(); go = os.path.join(t, "g.json")
        subprocess.run([os.path.abspath(exe), str(OUT / f"{arm}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + go], capture_output=True, timeout=600)
        g = json.load(open(go, encoding="utf-8"))["pages"][0]["glyphs"]; rows = {}
        for x in g:
            rows.setdefault(round(x["baseline"], 2), []).append(x["char"])
        body = [y for y, c in rows.items() if c and c[0] == "B"]
        ftr = [y for y, c in rows.items() if "".join(c).startswith("This")]
        res[arm] = {"n": len(body), "last": max(body) if body else None, "ftr": ftr[0] if ftr else None}
    summarize(res, "oxi ")
    wd = json.loads((OUT / "word.json").read_text(encoding="utf-8"))
    bad = [a for a in ARMS if res[a]["n"] != wd[a]["n"]]
    print("arms with a different page-1 count:", len(bad), bad[:12])


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
