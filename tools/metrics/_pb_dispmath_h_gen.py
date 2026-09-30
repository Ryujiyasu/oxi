# -*- coding: utf-8 -*-
"""S1613 probe: how tall is a display equation (oMathPara) line in Word?

Faithful slice of blind-G EN educational__005f2e39 (its styles / settings /
theme / sectPr, docDefaults after 200 + line 276).  Each arm is its own
document:  marker "M1" / EQUATION paragraph / marker "M2", markers with
spacing 0/0 single.  The equation paragraph keeps the document's pPr
(line 240 auto, ind 270, jc both, docDefaults after 200) unless the arm says
otherwise.  Equations:
  x     a lone x
  ab    a/b
  teq   the document's own t = [X1-X2]/(S sqrt(1/n1+1/n2))
  seq   the document's own S = sqrt(((n1-1)S1^2+(n2-1)S2^2)/(n1+n2-2))
Spacing variants:  _a200 (as the document)  _a0 (after 0)
Read: Word COM Information(6) of M1, EQ, M2 (collapsed starts) and the Word PDF
baselines of M1 / M2, so line height = M2 - M1 minus the marker line.

    python tools/metrics/_pb_dispmath_h_gen.py gen
    python tools/metrics/_pb_dispmath_h_gen.py word
    python tools/metrics/_pb_dispmath_h_gen.py oxi <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/educational/005f2e3927577fbe.docx"
OUT = REPO / "tests/fixtures/dispmath_h"
MARK = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
        '<w:t>{t}</w:t></w:r></w:p>')


def mr(t, sty=None):
    rpr = f'<m:rPr><m:sty m:val="{sty}"/></m:rPr>' if sty else ""
    return f"<m:r>{rpr}<m:t>{t}</m:t></m:r>"


def eqpara(inner, after):
    sp = '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>' if after == 0 else '<w:spacing w:line="240" w:lineRule="auto"/>'
    return (f'<w:p><w:pPr>{sp}<w:ind w:left="270"/><w:jc w:val="both"/></w:pPr>'
            f'<m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>{inner}</m:oMath></m:oMathPara></w:p>')


def doc_eqs(doc):
    out = []
    for m in re.finditer(r"<m:oMathPara>.*?</m:oMathPara>", doc, re.S):
        out.append(re.search(r"<m:oMath>(.*)</m:oMath>", m.group(0), re.S).group(1))
    return out


def arms(doc):
    eqs = doc_eqs(doc)
    teq = next(e for e in eqs if ">t<" in e)
    seq = next(e for e in eqs if "<m:rad>" in e and "S" in e and e is not teq)
    shapes = {
        "x": mr("x"),
        "ab": "<m:f><m:fPr/><m:num>" + mr("a") + "</m:num><m:den>" + mr("b") + "</m:den></m:f>",
        "teq": teq,
        "seq": seq,
    }
    return {f"{k}_a{a}": eqpara(v, a) for k, v in shapes.items() for a in (200, 0)}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    sect = re.sub(r'<w:cols [^>]*/>|<w:cols\b.*?</w:cols>', '<w:cols w:space="720"/>', sect, flags=re.S)
    for name, eq in arms(doc).items():
        body = MARK.format(t="M1") + eq + MARK.format(t="M2")
        xml = head + body + sect + "</w:body></w:document>"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"{name}.docx").write_bytes(buf.getvalue())
    print("ok")


def word():
    import win32com.client, fitz
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    res = {}
    try:
        for f in sorted(OUT.glob("*.docx")):
            tmp = os.path.join(tempfile.mkdtemp(), f.name)
            shutil.copy(f, tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            try:
                ys = []
                for i in (1, 2, 3):
                    r = d.Paragraphs(i).Range
                    ys.append(d.Range(r.Start, r.Start).Information(6))
                pdf = os.path.join(tempfile.mkdtemp(), "o.pdf")
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(0)
            base = {}
            for b in fitz.open(pdf)[0].get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    for s in l["spans"]:
                        if s["text"].strip() in ("M1", "M2"):
                            base[s["text"].strip()] = round(s["origin"][1], 2)
            res[f.stem] = {"info6": ys, "pdf": base,
                           "m2_m1": round(base["M2"] - base["M1"], 2) if len(base) == 2 else None}
            print(f.stem, res[f.stem], flush=True)
    finally:
        w.Quit()
    (OUT / "word.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


def oxi(exe):
    res = {}
    for f in sorted(OUT.glob("*.docx")):
        dump = OUT / f"_o_{f.stem}.json"
        subprocess.run([os.path.abspath(exe), str(f), str(OUT / "_o"), "--dump-layout=" + str(dump)], capture_output=True)
        d = json.load(open(dump, encoding="utf-8"))
        ys = {}
        for e in d["pages"][0]["elements"]:
            t = (e.get("text") or "").strip()
            if e.get("type") == "text" and t in ("M1", "M2"):
                ys[t] = round(e["y"], 2)
        res[f.stem] = {"m2_m1": round(ys["M2"] - ys["M1"], 2) if len(ys) == 2 else None}
        dump.unlink()
        print(f.stem, res[f.stem])
    for x in OUT.glob("_o*.png"):
        x.unlink()
    (OUT / "oxi.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
