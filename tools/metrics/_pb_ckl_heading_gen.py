# -*- coding: utf-8 -*-
"""S1614 probe: why is a ChecklistLevel1 heading row 16.4pt in Word?

Full-package controls of blind-G EN policies__0097fbf2 (one attribute each):
  A      original
  NONUM  ChecklistLevel1 heading paragraphs get <w:numPr><w:numId w:val="0"/></w:numPr>
  NOLDR  the heading paragraphs' mark rStyle ChecklistLeader removed
  ARIAL  numbering level 0 of numId 34's abstract: rFonts -> Arial
  NOSZ   the heading paragraphs' mark <w:sz w:val="20"/> removed (mark takes 12pt Leader)
Read (Word COM): Information(6) of the paragraph "Adequate provisions to solicit"
and of the next paragraph; their difference is the heading row pitch.

    python tools/metrics/_pb_ckl_heading_gen.py gen
    python tools/metrics/_pb_ckl_heading_gen.py word
    python tools/metrics/_pb_ckl_heading_gen.py oxi <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/policies/0097fbf26abcf61b.docx"
OUT = REPO / "tests/fixtures/ckl_heading"
KEY = "Adequate provisions to solicit"
NEXT = "Assent will be obtained from"
ARMS = ["A", "NONUM", "NOLDR", "ARIAL", "NOSZ"]


def heading_ppr_sub(doc, fn):
    # every paragraph whose pPr names ChecklistLevel1
    def rep(m):
        return fn(m.group(0))
    return re.sub(r'<w:pPr><w:pStyle w:val="ChecklistLevel1"/>.*?</w:pPr>', rep, doc, flags=re.S)


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    num = zin.read("word/numbering.xml").decode("utf-8")
    variants = {}
    variants["A"] = (doc, num)
    variants["NONUM"] = (heading_ppr_sub(doc, lambda p: p.replace('<w:pStyle w:val="ChecklistLevel1"/>',
                         '<w:pStyle w:val="ChecklistLevel1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>', 1)), num)
    variants["NOLDR"] = (heading_ppr_sub(doc, lambda p: p.replace('<w:rStyle w:val="ChecklistLeader"/>', '')), num)
    variants["NOSZ"] = (heading_ppr_sub(doc, lambda p: re.sub(r'(<w:rPr>.*?)<w:sz w:val="20"/>', r'\1', p, count=1, flags=re.S)), num)
    m = re.search(r'<w:lvl w:ilvl="0">(?:(?!</w:lvl>).)*?<w:pStyle w:val="ChecklistLevel1"/>.*?</w:lvl>', num, re.S)
    lvl = m.group(0)
    lvl2 = re.sub(r'<w:rFonts[^>]*/>', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial"/>', lvl)
    variants["ARIAL"] = (doc, num.replace(lvl, lvl2))
    for name in ARMS:
        d, n = variants[name]
        if name != "A":
            assert (d, n) != (doc, num), name
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = d.encode("utf-8")
                elif item.filename == "word/numbering.xml":
                    data = n.encode("utf-8")
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
                    for k in (KEY, NEXT):
                        if k in t and k not in ys:
                            c = d.Range(r.Start, r.Start)
                            ys[k] = (c.Information(3), c.Information(6))
                res[name] = {"head": ys.get(KEY), "next": ys.get(NEXT),
                             "pitch": round(ys[NEXT][1] - ys[KEY][1], 2) if len(ys) == 2 else None}
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
        ys = {}
        for pi, p in enumerate(d["pages"]):
            lines = {}
            for e in p["elements"]:
                if e.get("type") == "text":
                    lines.setdefault(round(e["y"], 2), []).append((e["x"], e["text"]))
            for y, v in sorted(lines.items()):
                t = "".join(s for _, s in sorted(v))
                for k in (KEY, NEXT):
                    if k in t and k not in ys:
                        ys[k] = (pi + 1, y)
        res[name] = {"pitch": round(ys[NEXT][1] - ys[KEY][1], 2) if len(ys) == 2 else None}
        dump.unlink()
        print(name, res[name])
    for x in OUT.glob("_o*.png"):
        x.unlink()
    (OUT / "oxi.json").write_text(json.dumps(res, indent=1), encoding="utf-8")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
