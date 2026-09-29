# -*- coding: utf-8 -*-
"""S1600 probe: why Word's heading rows in blind-G EN policies__0097fbf2 are
17.25pt where Oxi gives 12.7 (Calibri 10 + border).

Full-package controls (one edit per arm) of the row holding "Adequate provisions":
  A  original
  B  the endnoteReference run removed
  C  the checkbox cell's run font Segoe UI Symbol -> Calibri
  D  the endnoteReference run's rStyle removed (plain run, not superscript)
Measured with Word COM: Info(6) of the row's first paragraph and of the next row.

    python tools/metrics/_pb_enref_row_gen.py gen
    python tools/metrics/_pb_enref_row_gen.py measure
"""
import copy, io, json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/policies/0097fbf26abcf61b.docx"
OUT = REPO / "tests/fixtures/enref_row"
KEY = "Adequate provisions are made for soliciting the permission of parents or guardians"


def variants(doc):
    i = doc.index(KEY)
    t = doc.rindex("<w:tr ", 0, i)
    e = doc.index("</w:tr>", i) + len("</w:tr>")
    row = doc[t:e]
    ref = re.search(r'<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:endnoteReference[^>]*/></w:r>', row, re.S)
    assert ref
    b = row.replace(ref.group(0), "", 1)
    c = row.replace('w:ascii="Segoe UI Symbol" w:eastAsia="MS Gothic" w:hAnsi="Segoe UI Symbol" w:cs="Segoe UI Symbol"',
                    'w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"', 1)
    assert c != row
    d = row.replace(ref.group(0), ref.group(0).replace('<w:rStyle w:val="EndnoteReference"/>', ""), 1)
    assert d != row
    return {"A": doc, "B": doc[:t] + b + doc[e:], "C": doc[:t] + c + doc[e:], "D": doc[:t] + d + doc[e:]}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    for k, xml in variants(doc).items():
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"er_{k}.docx").write_bytes(buf.getvalue())
    print("ok")


def measure():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    res = {}
    try:
        for f in sorted(OUT.glob("er_*.docx")):
            tmp = os.path.join(tempfile.mkdtemp(), f.name)
            shutil.copy(f, tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            try:
                rng = d.Content
                fnd = rng.Find
                fnd.Execute(FindText="Adequate provisions are made for soliciting")
                r0 = d.Range(rng.Start, rng.Start)
                y0 = r0.Information(6)
                r2 = d.Content
                r2.Find.Execute(FindText="One of the following")
                y1 = d.Range(r2.Start, r2.Start).Information(6)
                res[f.stem] = {"row_top_text": y0, "next_row": y1, "delta": round(y1 - y0, 2)}
                print(f.stem, res[f.stem], flush=True)
            finally:
                d.Close(0)
    finally:
        w.Quit()
    (OUT / "enref_row_result.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


if __name__ == "__main__":
    {"gen": gen, "measure": measure}[sys.argv[1]]()
