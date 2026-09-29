# -*- coding: utf-8 -*-
"""S1606 probe: when does a one-line table row go to the next COLUMN whole?

Faithful slice of blind-G policies__1e87d3e6 (the vaccination table, 2-column
section).  Each arm = its own document:  exact spacer paragraph of height H,
then the original 18-row table, then the section's own sectPr.
Families:
  M  original rows (row 14 opens a vMerge over rows 14-15 in columns 0 and 2)
  P  vMerge removed from rows 14/15 (restart and continue cells become plain)
  S  as M but the section has ONE column (row 14 goes to page 2 instead)
  MV / PV  as M / P with every w:vAlign removed
  PB24  P with every border sz 4 -> 24 (0.5 -> 3pt)
  PM0 / PM200  P with tblCellMar bottom 43 -> 0 / 200 tw
H sweeps so the room left under row 13 crosses row 14's height.
Read (Word COM): x/page of row 14's middle-cell text and y of row 13's last line.

    python tools/metrics/_pb_vmerge_colbottom_gen.py gen
    python tools/metrics/_pb_vmerge_colbottom_gen.py word
    python tools/metrics/_pb_vmerge_colbottom_gen.py oxi <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/policies/1e87d3e6c31c432c.docx"
OUT = REPO / "tests/fixtures/vmerge_colbottom"
_hs = [float(x) for x in os.environ.get("VMC_HS", "252,0.5,29").split(",")]
HS = [_hs[0] + _hs[1] * i for i in range(int(_hs[2]))]
FAMS = [f for f in os.environ.get("VMC_FAMS", "M,P,S").split(",")]
ROW14 = "・65歳以上の人"
ROW13 = "(※８)"


def key(f, h):
    return f"{f}_{int(round(h * 100))}"


def slice_parts():
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    k = doc.index(ROW14)
    ts = doc.rindex("<w:tbl>", 0, k)
    te = doc.index("</w:tbl>", k) + len("</w:tbl>")
    sect = re.search(r"<w:sectPr\b.*?</w:sectPr>", doc[k:], re.S).group(0)
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    sect = re.sub(r"<w:type [^>]*/>", "", sect)
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    return zin, head, doc[ts:te], sect


def plain(tbl):
    rows = re.findall(r"<w:tr\b.*?</w:tr>", tbl, re.S)
    out = tbl
    for r in rows[14:16]:
        out = out.replace(r, re.sub(r'<w:vMerge(?: w:val="restart")?/>', "", r), 1)
    assert out != tbl
    return out


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin, head, tbl, sect = slice_parts()
    nov = lambda t: re.sub(r'<w:vAlign w:val="[a-z]+"/>', "", t)
    tbls = {"M": tbl, "P": plain(tbl), "S": tbl, "MV": nov(tbl), "PV": nov(plain(tbl))}
    # border width / bottom cell margin variants of P (one attribute each)
    mb = lambda t, v: t.replace('<w:bottom w:w="43" w:type="dxa"/>', f'<w:bottom w:w="{v}" w:type="dxa"/>', 1)
    tbls["PB24"] = plain(tbl).replace('w:sz="4"', 'w:sz="24"')
    tbls["PM0"] = mb(plain(tbl), 0)
    tbls["PM200"] = mb(plain(tbl), 200)
    one = re.sub(r'<w:cols w:num="2"[^>]*>.*?</w:cols>', '<w:cols w:space="720"/>', sect, count=1, flags=re.S)
    assert one != sect
    for f in FAMS:
        for h in HS:
            sp = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{round(h * 20)}" '
                  f'w:lineRule="exact"/><w:snapToGrid w:val="0"/></w:pPr><w:r><w:rPr><w:sz w:val="16"/></w:rPr>'
                  f'<w:t>S{key(f, h)}</w:t></w:r></w:p>')
            xml = head + sp + tbls[f] + "<w:p/>" + (one if f == "S" else sect) + "</w:body></w:document>"
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                    zout.writestr(copy.copy(item), data)
            (OUT / f"{key(f, h)}.docx").write_bytes(buf.getvalue())
    print("ok", len(FAMS) * len(HS))


def show(res):
    for f in FAMS:
        print(f, "".join({"R": "R", "L": ".", None: "?"}.get(res.get(key(f, h), {}).get("col")) for h in HS),
              f"(H {HS[0]}..{HS[-1]} step .5; R = row 14 in right column)")
        for h in HS:
            r = res.get(key(f, h), {})
            print("  ", h, r)


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    wj = OUT / "word.json"
    res = json.loads(wj.read_text(encoding="utf-8")) if wj.exists() else {}
    try:
        for f in FAMS:
            for h in HS:
                src = OUT / f"{key(f, h)}.docx"
                tmp = os.path.join(tempfile.mkdtemp(), src.name)
                shutil.copy(src, tmp)
                d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
                try:
                    d.Repaginate()
                    r14 = r13 = None
                    for i in range(1, d.Paragraphs.Count + 1):
                        rg = d.Paragraphs(i).Range
                        t = rg.Text
                        if ROW14 in t:
                            c = d.Range(rg.Start, rg.Start)
                            r14 = (c.Information(3), c.Information(5), c.Information(6))
                        elif ROW13 in t:
                            c = d.Range(rg.Start, rg.Start)
                            r13 = c.Information(6)
                    res[key(f, h)] = {"r14": r14, "r13_y": r13,
                                      "col": ("R" if r14 and (r14[0] > 1 or r14[1] > 300) else "L") if r14 else None}
                finally:
                    d.Close(0)
                print(key(f, h), res[key(f, h)], flush=True)
    finally:
        w.Quit()
    (OUT / "word.json").write_text(json.dumps(res, indent=0), encoding="utf-8")
    show(res)


def oxi(exe):
    oj = OUT / "oxi.json"
    res = json.loads(oj.read_text(encoding="utf-8")) if oj.exists() else {}
    for f in FAMS:
        for h in HS:
            src = OUT / f"{key(f, h)}.docx"
            dump = OUT / f"_o_{key(f, h)}.json"
            subprocess.run([os.path.abspath(exe), str(src), str(OUT / "_o"), "--dump-layout=" + str(dump)],
                           capture_output=True)
            dd = json.load(open(dump, encoding="utf-8"))
            r14 = r13 = None
            for pi, p in enumerate(dd["pages"]):
                for e in p["elements"]:
                    if e.get("type") != "text" or e.get("para_idx") != 1:
                        continue
                    if e.get("cell_row_idx") == 14 and e.get("cell_col_idx") == 1 and r14 is None:
                        r14 = (pi + 1, round(e["x"], 2), round(e["y"], 2))
                    elif e.get("cell_row_idx") == 13:
                        r13 = max(r13 or 0, round(e["y"], 2))
            res[key(f, h)] = {"r14": r14, "r13_y": r13,
                              "col": ("R" if r14 and (r14[0] > 1 or r14[1] > 300) else "L") if r14 else None}
            dump.unlink()
    for x in OUT.glob("_o*.png"):
        x.unlink()
    (OUT / "oxi.json").write_text(json.dumps(res, indent=0), encoding="utf-8")
    show(res)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
