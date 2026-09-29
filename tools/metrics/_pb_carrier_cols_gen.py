# -*- coding: utf-8 -*-
"""S1594 probe: does an empty continuous section-break carrier that follows a
table keep its line when the NEXT section changes the column count?

Full-package controls of blind-G policies__1e87d3e6 (only one attribute edited
per arm, raw-byte edits so namespaces stay as written):
  A  original (carrier Heading-1 style, next section 2 columns)
  B  next section's <w:cols w:num="2"...> -> one column (num attribute removed)
  C  carrier's pStyle removed (Normal)
  D  B + C
Measured with Word COM: page / Information(6) of the carrier, the exact-1pt
paragraph after it, and the heading.

    python tools/metrics/_pb_carrier_cols_gen.py gen
    python tools/metrics/_pb_carrier_cols_gen.py measure
"""
import copy, io, json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/policies/1e87d3e6c31c432c.docx"
OUT = REPO / "tests/fixtures/carrier_cols"
HEAD = "都営交通の無料乗車券"


def variants(doc):
    k = doc.index(HEAD)
    t = doc.rindex("</w:tbl>", 0, k)
    carrier_s = t + len("</w:tbl>")
    carrier_e = doc.index("</w:p>", carrier_s) + len("</w:p>")
    carrier = doc[carrier_s:carrier_e]
    nxt = re.search(r"<w:sectPr\b.*?</w:sectPr>", doc[k:], re.S)
    ns, ne = k + nxt.start(), k + nxt.end()
    sect = doc[ns:ne]
    one_col = re.sub(r'<w:cols w:num="2"([^>]*)>.*?</w:cols>', r'<w:cols\1/>', sect, count=1, flags=re.S)
    assert one_col != sect
    plain = carrier.replace('<w:pStyle w:val="1"/>', "", 1)
    assert plain != carrier

    def build(c, s):
        return doc[:carrier_s] + c + doc[carrier_e:ns] + s + doc[ne:]
    ex20 = carrier.replace("<w:pPr>", '<w:pPr><w:spacing w:line="20" w:lineRule="exact"/>', 1)
    ex700 = carrier.replace("<w:pPr>", '<w:pPr><w:spacing w:line="700" w:lineRule="exact"/>', 1)
    return {"A": doc, "B": build(carrier, one_col), "C": build(plain, sect), "D": build(plain, one_col),
            "E": build(ex20, sect), "F": build(ex700, sect),
            "G": build(carrier.replace("<w:pPr>", '<w:pPr><w:spacing w:line="300" w:lineRule="exact"/>', 1), sect),
            "H": build(carrier.replace("<w:pPr>", '<w:pPr><w:spacing w:line="400" w:lineRule="exact"/>', 1), sect)}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    for name, xml in variants(doc).items():
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"cc_{name}.docx").write_bytes(buf.getvalue())
        print("wrote", name)


def measure():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    res = {}
    try:
        for f in sorted(OUT.glob("cc_*.docx")):
            tmp = os.path.join(tempfile.mkdtemp(), f.name)
            shutil.copy(f, tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            try:
                d.Repaginate()
                n = d.Paragraphs.Count
                hi = None
                for i in range(1, n + 1):
                    if HEAD in d.Paragraphs(i).Range.Text:
                        hi = i
                        break
                rows = []
                for i in range(hi - 2, hi + 1):
                    r = d.Paragraphs(i).Range
                    c = d.Range(r.Start, r.Start)
                    rows.append({"i": i, "page": c.Information(3), "y": c.Information(6),
                                 "text": r.Text[:12]})
                res[f.stem] = {"pages": d.ComputeStatistics(2), "rows": rows}
                print(f.stem, res[f.stem], flush=True)
            finally:
                d.Close(0)
    finally:
        w.Quit()
    (OUT / "carrier_cols_result.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "measure": measure}[sys.argv[1]]()
