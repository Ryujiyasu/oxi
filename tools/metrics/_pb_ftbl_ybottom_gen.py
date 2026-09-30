# -*- coding: utf-8 -*-
"""S1615 probe: a floating table with vertAnchor=text + tblpYSpec=bottom.

Full-package controls of blind-G JA forms__02157d72 (one attribute each) on the
緊急連絡先 table (tblpXSpec=center, tblpYSpec=bottom, tblOverlap never), which
directly follows the paragraph «※電話番号に変更…»:
  A       original
  Y0      tblpYSpec removed, tblpY="0"
  Y35     tblpYSpec removed, tblpY="35"
  YTOP    tblpYSpec="top"
  OVL     tblOverlap removed
  NOTEAFT the ※ paragraph moved to just after the table
  INLINE  tblpPr removed (an inline table)
  EMPTY   the ※ paragraph's runs removed
Read (Word COM): page / Information(6) of ※, of the table's first cell and of
the paragraph after the table («１．かかりつけ医…» two paragraphs later).

    python tools/metrics/_pb_ftbl_ybottom_gen.py gen
    python tools/metrics/_pb_ftbl_ybottom_gen.py word
    python tools/metrics/_pb_ftbl_ybottom_gen.py oxi <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/forms/02157d72758f9c93.docx"
OUT = REPO / "tests/fixtures/ftbl_ybottom"
NOTE = "電話番号に変更"
CELL = "緊急連絡先"
AFTER = "かかりつけ医がある場合"
ARMS = [x for x in os.environ.get("FTB_ARMS", "A,Y0,Y35,YTOP,OVL,NOTEAFT,INLINE,EMPTY").split(",")]


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    k = doc.index('w:tblpYSpec="bottom"')
    ts = doc.rindex("<w:tbl>", 0, k)
    te = doc.index("</w:tbl>", k) + len("</w:tbl>")
    tbl = doc[ts:te]
    pp = re.search(r'<w:tblpPr[^>]*/>', tbl).group(0)
    ne = doc.rindex("</w:p>", 0, ts) + len("</w:p>")
    ns = max(doc.rindex("<w:p ", 0, ne), doc.rindex("<w:p>", 0, ne) if "<w:p>" in doc[:ne] else -1)
    note = doc[ns:ne]
    assert ne <= ts
    v = {}
    v["A"] = doc
    v["Y0"] = doc.replace(pp, pp.replace('w:tblpYSpec="bottom"', 'w:tblpY="0"'), 1)
    v["Y35"] = doc.replace(pp, pp.replace('w:tblpYSpec="bottom"', 'w:tblpY="35"'), 1)
    v["YTOP"] = doc.replace(pp, pp.replace('w:tblpYSpec="bottom"', 'w:tblpYSpec="top"'), 1)
    v["OVL"] = doc[:ts] + tbl.replace('<w:tblOverlap w:val="never"/>', '', 1) + doc[te:]
    v["NOTEAFT"] = doc[:ns] + doc[ne:te] + note + doc[te:]
    v["INLINE"] = doc[:ts] + tbl.replace(pp, "", 1).replace('<w:tblOverlap w:val="never"/>', '', 1) + doc[te:]
    # the note paragraph with its runs removed (an empty anchor paragraph)
    empty_note = re.sub(r"<w:r\b[^>]*>.*?</w:r>", "", note, flags=re.S)
    v["EMPTY"] = doc[:ns] + empty_note + doc[ne:]
    for name in ARMS:
        if name != "A":
            assert v[name] != doc, name
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = v[name].encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"{name}.docx").write_bytes(buf.getvalue())
    print("ok")


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    res = {}
    try:
        for name in ARMS:
            tmp = os.path.join(tempfile.mkdtemp(), name + ".docx")
            shutil.copy(OUT / f"{name}.docx", tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            try:
                ys = {}
                for i in range(1, d.Paragraphs.Count + 1):
                    r = d.Paragraphs(i).Range
                    t = r.Text
                    for k in (NOTE, CELL, AFTER):
                        if k in t and k not in ys:
                            c = d.Range(r.Start, r.Start)
                            ys[k] = (c.Information(3), c.Information(6))
                res[name] = {"note": ys.get(NOTE), "cell": ys.get(CELL), "after": ys.get(AFTER)}
            finally:
                d.Close(0)
            print(name, res[name], flush=True)
    finally:
        w.Quit()
    (OUT / "word.json").write_text(json.dumps(res, indent=1, ensure_ascii=False), encoding="utf-8")


def oxi(exe):
    res = {}
    for name in ARMS:
        dump = OUT / f"_o_{name}.json"
        subprocess.run([os.path.abspath(exe), str(OUT / f"{name}.docx"), str(OUT / "_o"), "--dump-layout=" + str(dump)],
                       capture_output=True)
        d = json.load(open(dump, encoding="utf-8"))
        ys = {}
        for pi, p in enumerate(d["pages"]):
            lines = {}
            for e in p["elements"]:
                if e.get("type") == "text":
                    lines.setdefault(round(e["y"], 2), []).append((e["x"], e["text"]))
            for y, parts in sorted(lines.items()):
                t = "".join(s for _, s in sorted(parts))
                for k in (NOTE, CELL, AFTER):
                    if k in t and k not in ys:
                        ys[k] = (pi + 1, y)
        res[name] = {"note": ys.get(NOTE), "cell": ys.get(CELL), "after": ys.get(AFTER)}
        dump.unlink()
        print(name, res[name])
    for x in OUT.glob("_o*.png"):
        x.unlink()
    (OUT / "oxi.json").write_text(json.dumps(res, indent=1, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
