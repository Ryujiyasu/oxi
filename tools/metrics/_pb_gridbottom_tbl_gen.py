# -*- coding: utf-8 -*-
"""S1617 probe: typed-grid page bottom -- full line box or natural height?

Faithful slice of JA policies__074da728 (lines grid 334 = 16.7pt, MS Mincho 12,
A4, bottom margin 1418): exact spacer H, then blocks 18..23 of the document
(【嘱託職員】 / the 3-row table / three notes / «３　勤務場所»).  Families:
  T   as the document (table present)
  N   the table replaced by three plain paragraphs of its first cell's text
      (same flow without a table)
  S   as N plus a one-line one-cell table after the plain paragraphs
  U   as N with the original table moved to the TOP (before the spacer)
H sweeps so the heading's line crosses the page bottom.
Read: page of «３　勤務場所» (Word COM) and, for Oxi, the heading's line box.

    python tools/metrics/_pb_gridbottom_tbl_gen.py gen
    python tools/metrics/_pb_gridbottom_tbl_gen.py word
    python tools/metrics/_pb_gridbottom_tbl_gen.py oxi <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/policies/074da7283a735cb5.docx"
OUT = REPO / "tests/fixtures/gridbottom_tbl"
HEAD = "勤務場所"
_hs = [float(x) for x in os.environ.get("GBT_HS", "440,1,40").split(",")]
HS = [_hs[0] + _hs[1] * i for i in range(int(_hs[2]))]
FAMS = os.environ.get("GBT_FAMS", "T,N").split(",")


def key(f, h):
    return f"{f}_{int(round(h * 100))}"


def blocks(body):
    out = []
    i = 0
    pat = re.compile(r"<w:(p|tbl)\b")
    while True:
        m = pat.search(body, i)
        if not m:
            break
        if m.group(1) == "tbl":
            e = body.index("</w:tbl>", m.start()) + len("</w:tbl>")
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
    seg = bl[18:24]
    assert seg[1][0] == "tbl" and HEAD in re.sub(r"<[^>]+>", "", seg[5][1]), [re.sub(r"<[^>]+>", "", x[1])[:10] for x in seg]
    tbl_text = re.sub(r"<[^>]+>", "", seg[1][1])
    plain = "".join(f'<w:p><w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
                    f'<w:sz w:val="24"/></w:rPr><w:t>{t}</w:t></w:r></w:p>'
                    for t in ("業務の概要", "地域住民の複合・複雑化した支援", "ズに対応する包括的な支援"))
    for f in FAMS:
        for h in HS:
            sp = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{round(h * 20)}" w:lineRule="exact"/>'
                  f'</w:pPr><w:r><w:t>S{key(f, h)}</w:t></w:r></w:p>')
            small = ('<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/>'
                     '<w:bottom w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>'
                     '<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
            if f == "T":
                pre, mid = "", seg[1][1]
            elif f == "S":
                pre, mid = "", plain + small
            elif f == "U":
                pre, mid = seg[1][1], plain
            else:
                pre, mid = "", plain
            xml = head + pre + sp + seg[0][1] + mid + "".join(x[1] for x in seg[2:]) + "<w:p/>" + sect + "</w:body></w:document>"
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
