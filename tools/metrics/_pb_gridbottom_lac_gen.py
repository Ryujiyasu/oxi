# -*- coding: utf-8 -*-
"""S1619 probe: linesAndChars page bottom under a table -- where does Word flip?

Faithful slice of JA policies__1f014c0f (linesAndChars 416 = 20.8pt, HGP
Gothic M 12, A4, bottom margin 1134): exact spacer H, then blocks 372..375 of
the document («４　記録» / the 2-row 記録 table / «※帳票は　　年保存する。» / the
spaces paragraph).  Families:
  T   as the document (table present)
  H   the table removed (heading, then the ※ line)
H sweeps so the ※ line crosses the page bottom.
Read: page and Info(6) of «※帳票は» (Word COM) and, for Oxi, its line box.

    python tools/metrics/_pb_gridbottom_lac_gen.py gen
    python tools/metrics/_pb_gridbottom_lac_gen.py word
    python tools/metrics/_pb_gridbottom_lac_gen.py oxi <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/policies/1f014c0fdd5ce4e1.docx"
OUT = REPO / "tests/fixtures/gridbottom_lac"
HEAD = "帳票は"
_hs = [float(x) for x in os.environ.get("GBL_HS", "618,0.5,33").split(",")]
HS = [_hs[0] + _hs[1] * i for i in range(int(_hs[2]))]
FAMS = os.environ.get("GBL_FAMS", "T,H").split(",")


def key(f, h):
    return f"{f}_{int(round(h * 100))}"


def blocks(body):
    """Top-level <w:p>/<w:tbl> blocks; a table's end is found by nesting depth."""
    out = []
    i = 0
    pat = re.compile(r"<w:(p|tbl)\b")
    while True:
        m = pat.search(body, i)
        if not m:
            break
        if m.group(1) == "tbl":
            depth = 0
            for mm in re.finditer(r"<w:tbl[ >]|</w:tbl>", body[m.start():]):
                depth += -1 if mm.group(0) == "</w:tbl>" else 1
                if depth == 0:
                    e = m.start() + mm.end()
                    break
        else:
            e = body.index("</w:p>", m.start()) + len("</w:p>")
        out.append((m.group(1), body[m.start():e]))
        i = e
    return out


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    b0 = doc.index("<w:body>") + len("<w:body>")
    head = doc[:b0]
    body = doc[b0:]
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    bl = blocks(body)
    seg = bl[372:376]
    txt = [re.sub(r"<[^>]+>", "", x[1])[:12] for x in seg]
    assert seg[1][0] == "tbl" and HEAD in txt[2] and "記録" in txt[0], txt
    for f in FAMS:
        for h in HS:
            sp = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{round(h * 20)}" w:lineRule="exact"/>'
                  f'</w:pPr><w:r><w:t>S{key(f, h)}</w:t></w:r></w:p>')
            mid = seg[1][1] if f == "T" else ""
            xml = head + sp + seg[0][1] + mid + seg[2][1] + seg[3][1] + "<w:p/>" + sect + "</w:body></w:document>"
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                    zout.writestr(copy.copy(item), data)
            (OUT / f"{key(f, h)}.docx").write_bytes(buf.getvalue())
    print("ok", len(FAMS) * len(HS))


def show(res):
    for f in FAMS:
        print(f, "".join({1: ".", 2: "N"}.get((res.get(key(f, h)) or {}).get("page"), "?") for h in HS),
              f"(H {HS[0]}..{HS[-1]}; N = heading on page 2)")


def load(n):
    p = OUT / n
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
                tmp = os.path.join(tempfile.mkdtemp(), key(f, h) + ".docx")
                shutil.copy(OUT / f"{key(f, h)}.docx", tmp)
                d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
                try:
                    r = None
                    for i in range(1, d.Paragraphs.Count + 1):
                        rg = d.Paragraphs(i).Range
                        if HEAD in rg.Text:
                            c = d.Range(rg.Start, rg.Start)
                            r = {"page": c.Information(3), "y": c.Information(6)}
                            break
                    res[key(f, h)] = r
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
            dump = OUT / f"_o_{key(f, h)}.json"
            subprocess.run([os.path.abspath(exe), str(OUT / f"{key(f, h)}.docx"), str(OUT / "_o"), "--dump-layout=" + str(dump)],
                           capture_output=True)
            d = json.load(open(dump, encoding="utf-8"))
            r = None
            for pi, p in enumerate(d["pages"]):
                lines = {}
                for e in p["elements"]:
                    if e.get("type") == "text":
                        lines.setdefault(round(e["y"], 2), []).append((e["x"], e["text"], e["h"]))
                for y, parts in sorted(lines.items()):
                    t = "".join(s for _, s, _ in sorted(parts))
                    if HEAD in t and r is None:
                        r = {"page": pi + 1, "y": y, "bottom": round(y + max(hh for _, _, hh in parts), 2)}
            res[key(f, h)] = r
            dump.unlink()
    for x in OUT.glob("_o*.png"):
        x.unlink()
    (OUT / "oxi.json").write_text(json.dumps(res, indent=0), encoding="utf-8")
    show(res)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
