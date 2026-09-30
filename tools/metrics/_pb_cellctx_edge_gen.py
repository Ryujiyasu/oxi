# -*- coding: utf-8 -*-
"""S1612 probe: does contextualSpacing drop a table cell's outer spacing?

Faithful slice of blind-G EN educational__005f2e39 (its styles / settings /
theme, no compatibilityMode): the "Tabel 1.4" table alone, 7 rows, every cell
one paragraph TNR 9pt b/i with spacing before 60 / after 60 / line 240 auto and
contextualSpacing.  Arms (one attribute each):
  A     original
  NOCTX contextualSpacing removed from every cell paragraph
  B0    before=0 (ctx kept)
  A0    after=0 (ctx kept)
  C15   settings get compatibilityMode 15
  C15N  C15 + NOCTX
  STY / STYN  every cell paragraph in Style1 (ctx kept / removed)
Read: Word COM Information(6) of each row's first cell; row pitch = mean step
over rows 2..6.

    python tools/metrics/_pb_cellctx_edge_gen.py gen
    python tools/metrics/_pb_cellctx_edge_gen.py word
    python tools/metrics/_pb_cellctx_edge_gen.py oxi <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/educational/005f2e3927577fbe.docx"
OUT = REPO / "tests/fixtures/cellctx_edge"
ANCHOR = ">Lembang<"
ARMS = ["A", "NOCTX", "B0", "A0", "C15", "C15N", "STY", "STYN"]
C15 = ('<w:compat><w:compatSetting w:name="compatibilityMode" '
       'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    settings = zin.read("word/settings.xml").decode("utf-8")
    k = doc.index(ANCHOR)
    ts = doc.rindex("<w:tbl>", 0, k)
    te = doc.index("</w:tbl>", k) + len("</w:tbl>")
    tbl = doc[ts:te]
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    sect = re.sub(r'<w:cols [^>]*/>|<w:cols\b.*?</w:cols>', '<w:cols w:space="720"/>', sect, flags=re.S)
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    variants = {
        "A": tbl,
        "NOCTX": tbl.replace("<w:contextualSpacing/>", ""),
        "B0": tbl.replace('w:before="60"', 'w:before="0"'),
        "A0": tbl.replace('w:after="60"', 'w:after="0"'),
        "C15": tbl,
        "C15N": tbl.replace("<w:contextualSpacing/>", ""),
        # every cell paragraph in Style1 (not the cell-end mark's Normal)
        "STY": re.sub(r"(<w:p\b[^>]*><w:pPr>)", r'\1<w:pStyle w:val="Style1"/>', tbl),
        "STYN": re.sub(r"(<w:p\b[^>]*><w:pPr>)", r'\1<w:pStyle w:val="Style1"/>', tbl).replace("<w:contextualSpacing/>", ""),
    }
    for name in ARMS:
        body = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1440" w:lineRule="exact"/></w:pPr>'
                f'<w:r><w:t>S{name}</w:t></w:r></w:p>' + variants[name] + "<w:p/>")
        xml = head + body + sect + "</w:body></w:document>"
        st = settings
        if name.startswith("C15"):
            st = re.sub(r"<w:compat>.*?</w:compat>|<w:compat/>", "", st, flags=re.S)
            st = st.replace("</w:settings>", C15 + "</w:settings>")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = xml.encode("utf-8")
                elif item.filename == "word/settings.xml":
                    data = st.encode("utf-8")
                zout.writestr(copy.copy(item), data)
        (OUT / f"{name}.docx").write_bytes(buf.getvalue())
    print("ok", len(ARMS))


def pitch(ys):
    steps = [b - a for a, b in zip(ys[2:], ys[3:])]
    return round(sum(steps) / len(steps), 3) if steps else None


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
                t = d.Tables(1)
                first = {}
                for c in t.Range.Cells:
                    if c.RowIndex not in first:
                        r = c.Range
                        first[c.RowIndex] = d.Range(r.Start, r.Start).Information(6)
                ys = [first[k] for k in sorted(first)]
                res[name] = {"rows": ys, "pitch": pitch(ys)}
            finally:
                d.Close(0)
            print(name, res[name], flush=True)
    finally:
        w.Quit()
    (OUT / "word.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


def oxi(exe):
    res = {}
    for name in ARMS:
        dump = OUT / f"_o_{name}.json"
        subprocess.run([os.path.abspath(exe), str(OUT / f"{name}.docx"), str(OUT / "_o"), "--dump-layout=" + str(dump)],
                       capture_output=True)
        d = json.load(open(dump, encoding="utf-8"))
        first = {}
        for e in d["pages"][0]["elements"]:
            if e.get("type") == "text" and e.get("cell_row_idx") is not None and e.get("cell_col_idx") == 0:
                first.setdefault(e["cell_row_idx"], round(e["y"], 2))
        ys = [first[k] for k in sorted(first)]
        res[name] = {"rows": ys, "pitch": pitch(ys)}
        dump.unlink()
        print(name, res[name])
    for x in OUT.glob("_o*.png"):
        x.unlink()
    (OUT / "oxi.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
