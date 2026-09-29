# -*- coding: utf-8 -*-
"""S1606 probe (grid lines): what does a split row reserve under its last kept
line when the cell lines are docGrid-snapped auto lines?

Faithful slice of golden tokyoshugyo_000599795 (docGrid lines 360, MS Mincho
10.5, one-row table with 0.5pt borders, no tblCellMar).  Each arm: an exact
spacer of height H, then the original one-row table ("（時間外及び休日労働等）"),
then the section's sectPr.  Families:
  G      original (cell bottom margin 0)
  GM200  tblCellMar bottom 200 tw added
Read: the page of the line starting "れを所轄" (Word: collapsed range at that
character, Information(3); Oxi: the text element that starts with it) and the
y of that line.

    python tools/metrics/_pb_gridsplit_reserve_gen.py gen
    python tools/metrics/_pb_gridsplit_reserve_gen.py word
    python tools/metrics/_pb_gridsplit_reserve_gen.py oxi <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "tools/golden-test/documents/docx/tokyoshugyo_000599795.docx"
OUT = REPO / "tests/fixtures/gridsplit_reserve"
_hs = [float(x) for x in os.environ.get("GSR_HS", "560,0.5,41").split(",")]
HS = [_hs[0] + _hs[1] * i for i in range(int(_hs[2]))]
FAMS = os.environ.get("GSR_FAMS", "G,GM200").split(",")
HEAD = "（時間外及び休日労働等）"
LINE = "所轄"


def key(f, h):
    return f"{f}_{int(round(h * 100))}"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    k = doc.index(HEAD)
    ts = doc.rindex("<w:tbl>", 0, k)
    te = doc.index("</w:tbl>", k) + len("</w:tbl>")
    tbl = doc[ts:te]
    sect = re.search(r"<w:sectPr\b.*?</w:sectPr>", doc[k:], re.S).group(0)
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    sect = re.sub(r"<w:type [^>]*/>", "", sect)
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    m200 = tbl.replace("</w:tblBorders>", '</w:tblBorders><w:tblCellMar><w:bottom w:w="200" w:type="dxa"/></w:tblCellMar>', 1)
    assert m200 != tbl
    tbls = {"G": tbl, "GM200": m200}
    for f in FAMS:
        for h in HS:
            sp = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{round(h * 20)}" '
                  f'w:lineRule="exact"/><w:snapToGrid w:val="0"/></w:pPr><w:r><w:rPr><w:sz w:val="16"/></w:rPr>'
                  f'<w:t>S{key(f, h)}</w:t></w:r></w:p>')
            xml = head + sp + tbls[f] + "<w:p/>" + sect + "</w:body></w:document>"
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                    zout.writestr(copy.copy(item), data)
            (OUT / f"{key(f, h)}.docx").write_bytes(buf.getvalue())
    print("ok", len(FAMS) * len(HS))


def show(res):
    for f in FAMS:
        print(f.ljust(6), "".join({1: ".", 2: "N"}.get((res.get(key(f, h)) or {}).get("page"), "?") for h in HS),
              f"(H {HS[0]}..{HS[-1]}; N = line on page 2)")


def load(name):
    p = OUT / name
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    res = load("word.json")
    try:
        for f in FAMS:
            for h in HS:
                src = OUT / f"{key(f, h)}.docx"
                tmp = os.path.join(tempfile.mkdtemp(), src.name)
                shutil.copy(src, tmp)
                d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
                try:
                    d.Repaginate()
                    rg = d.Content
                    fnd = rg.Find
                    ok = fnd.Execute(LINE)
                    if ok:
                        c = d.Range(rg.Start, rg.Start)
                        res[key(f, h)] = {"page": c.Information(3), "y": c.Information(6)}
                    else:
                        res[key(f, h)] = None
                finally:
                    d.Close(0)
                print(key(f, h), res[key(f, h)], flush=True)
    finally:
        w.Quit()
    (OUT / "word.json").write_text(json.dumps(res, indent=0), encoding="utf-8")
    show(res)


def oxi(exe):
    res = load("oxi.json")
    for f in FAMS:
        for h in HS:
            src = OUT / f"{key(f, h)}.docx"
            dump = OUT / f"_o_{key(f, h)}.json"
            subprocess.run([os.path.abspath(exe), str(src), str(OUT / "_o"), "--dump-layout=" + str(dump)],
                           capture_output=True)
            dd = json.load(open(dump, encoding="utf-8"))
            r = None
            for pi, p in enumerate(dd["pages"]):
                for e in p["elements"]:
                    if e.get("type") == "text" and LINE in (e.get("text") or "") and r is None:
                        r = {"page": pi + 1, "y": round(e["y"], 2), "h": round(e["h"], 2)}
            res[key(f, h)] = r
            dump.unlink()
    for x in OUT.glob("_o*.png"):
        x.unlink()
    (OUT / "oxi.json").write_text(json.dumps(res, indent=0), encoding="utf-8")
    show(res)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
