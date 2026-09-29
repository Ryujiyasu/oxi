# -*- coding: utf-8 -*-
"""S1605 probe: does a table AFTER a typed-grid line change the page-bottom fit?

Faithful slice of blind-G policies__1f014c0f (linesAndChars, linePitch 416), one
arm per page:  exact spacer H  +  test paragraph  +  FOLLOWER  +  page break.
  B   follower = a body paragraph
  T1  follower = a one-row table (test paragraph is one line)
  T2  follower = a one-row table (test paragraph wraps to two lines; the spacer
      is one grid line shorter so the LAST line sits where T1's line sits)
H sweeps 705.0 .. 712.0 in 0.5pt steps.
Read: Word COM page of the test paragraph's LAST line vs its spacer's page.

    python tools/metrics/_pb_s603_gen.py gen
    python tools/metrics/_pb_s603_gen.py word
    python tools/metrics/_pb_s603_gen.py oxi <renderer.exe>
"""
import copy, io, json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/policies/1f014c0fdd5ce4e1.docx"
OUT = REPO / "tests/fixtures/s603"
HS = [705 + 0.5 * i for i in range(15)]
KINDS = ["B", "T1", "T2"]
RPR = ('<w:rPr><w:rFonts w:ascii="HGPｺﾞｼｯｸM" w:eastAsia="HGPｺﾞｼｯｸM" w:hint="eastAsia"/>'
       '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>')


def run(t):
    return f"<w:r>{RPR}<w:t>{t}</w:t></w:r>"


def table(tag):
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="108" w:type="dxa"/>'
            '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>'
            f'<w:p>{run("表" + tag)}</w:p></w:tc></w:tr></w:tbl>')


def key(kind, h):
    return f"{kind}_{int(h * 10)}"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    body = ""
    arms = [(k, h) for k in KINDS for h in HS]
    for i, (k, h) in enumerate(arms):
        tag = key(k, h)
        sp = h - (20.8 if k == "T2" else 0)
        body += f'<w:p><w:pPr><w:spacing w:line="{round(sp * 20)}" w:lineRule="exact"/></w:pPr>{run("S" + tag)}</w:p>'
        text = "※帳票は　　年保存する。" if k != "T2" else "※帳票は　　年保存する。" * 4
        body += f'<w:p><w:pPr><w:jc w:val="left"/></w:pPr>{run(text + "E" + tag)}</w:p>'
        body += table(tag) if k != "B" else f'<w:p>{run("本文F" + tag)}</w:p>'
        if i + 1 < len(arms):
            body += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    xml = head + body + sect + "</w:body></w:document>"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(copy.copy(item), data)
    (OUT / "s603.docx").write_bytes(buf.getvalue())
    print("ok", len(arms), "arms")


def show(res):
    for k in KINDS:
        ks = [res.get(key(k, h)) for h in HS]
        print(k.ljust(3), "".join("K" if x else "." for x in ks), f"(H {HS[0]}..{HS[-1]} step .5)")


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    res = {}
    try:
        src = OUT / "s603.docx"
        tmp = os.path.join(tempfile.mkdtemp(), src.name)
        shutil.copy(src, tmp)
        d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
        try:
            pg = {}
            for i in range(1, d.Paragraphs.Count + 1):
                r = d.Paragraphs(i).Range
                m = re.search(r"([SE])([A-Z0-9]+_\d+)", r.Text)
                if m:
                    pos = r.End - 1 if m.group(1) == "E" else r.Start
                    pg[m.group(1) + m.group(2)] = d.Range(pos, pos).Information(3)
            for k in KINDS:
                for h in HS:
                    t = key(k, h)
                    res[t] = pg.get("E" + t) == pg.get("S" + t)
        finally:
            d.Close(0)
    finally:
        w.Quit()
    (OUT / "word.json").write_text(json.dumps(res, indent=0), encoding="utf-8")
    show(res)


def oxi(exe):
    sys.path.insert(0, str(REPO / "tools" / "metrics"))
    os.environ["OXI_GDI_EXE"] = os.path.abspath(exe)
    import measure_pagination_oxi as MO
    MO.RENDERER = os.path.abspath(exe)
    import subprocess
    dump = OUT / "oxi_dump.json"
    subprocess.run([MO.RENDERER, str(OUT / "s603.docx"), str(OUT / "o"), "--dump-layout=" + str(dump)], capture_output=True)
    for f in OUT.glob("o*.png"):
        f.unlink()
    d = json.load(open(dump, encoding="utf-8"))
    last = {}
    for pi, p in enumerate(d["pages"]):
        for e in p["elements"]:
            t = e.get("text") or ""
            for m in re.finditer(r"([SE])([A-Z0-9]+_\d+)", t):
                last[m.group(1) + m.group(2)] = pi
    res = {}
    for k in KINDS:
        for h in HS:
            t = key(k, h)
            res[t] = last.get("E" + t) == last.get("S" + t)
    (OUT / "oxi.json").write_text(json.dumps(res, indent=0), encoding="utf-8")
    show(res)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "word": word, "oxi": lambda: oxi(sys.argv[2])}[sys.argv[1]]()
