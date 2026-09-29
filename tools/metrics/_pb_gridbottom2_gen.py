# -*- coding: utf-8 -*-
"""S1604 probe: when does a typed-grid line still fit at the page bottom?

Faithful slice of blind-G policies__1f014c0f (styles, settings, theme and the
linesAndChars / linePitch 416 section kept; footer references dropped). Arm k is
one page:  exact spacer of H_k pt  +  test paragraph (the document's own note
line, HGPGothicM 12pt)  +  page break.  H sweeps 670..730 pt in 1pt steps.
Read: Word COM page of each test paragraph (kept = same page as its spacer).

    python tools/metrics/_pb_gridbottom2_gen.py gen
    python tools/metrics/_pb_gridbottom2_gen.py word
    python tools/metrics/_pb_gridbottom2_gen.py oxi <renderer.exe>
"""
import copy, io, json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/policies/1f014c0fdd5ce4e1.docx"
OUT = REPO / "tests/fixtures/gridbottom"
HS = list(range(670, 731)) if "--fine" not in sys.argv else [14200 + i for i in range(0, 21)]  # fine: twips/20 steps 710.00..711.00
RUN = ('<w:r><w:rPr><w:rFonts w:ascii="HGPｺﾞｼｯｸM" w:eastAsia="HGPｺﾞｼｯｸM" w:hint="eastAsia"/>'
       '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>{t}</w:t></w:r>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    body = ""
    for i, h in enumerate(HS):
        tw = h if h > 10000 else h * 20
        spacer = f'<w:p><w:pPr><w:spacing w:line="{tw}" w:lineRule="exact"/></w:pPr>{RUN.format(t="S" + str(h))}</w:p>'
        test = f'<w:p><w:pPr><w:jc w:val="left"/></w:pPr>{RUN.format(t="※帳票は　　年保存する。T" + str(h))}</w:p>'
        brk = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' if i + 1 < len(HS) else ""
        body += spacer + test + brk
    xml = head + body + sect + "</w:body></w:document>"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(copy.copy(item), data)
    (OUT / ("gridbottom_fine.docx" if "--fine" in sys.argv else "gridbottom.docx")).write_bytes(buf.getvalue())
    print("ok", len(HS), "arms")


def kept(pages_of):
    # pages_of: {"S670": p, "T670": p, ...}
    return {h: pages_of.get(f"T{h}") == pages_of.get(f"S{h}") for h in HS}


def word():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    try:
        src = OUT / ("gridbottom_fine.docx" if "--fine" in sys.argv else "gridbottom.docx")
        tmp = os.path.join(tempfile.mkdtemp(), src.name)
        shutil.copy(src, tmp)
        d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
        try:
            pg, ys = {}, {}
            for i in range(1, d.Paragraphs.Count + 1):
                r = d.Paragraphs(i).Range
                t = r.Text.strip()
                m = re.search(r"([ST])(\d+)$", t)
                if m:
                    c = d.Range(r.Start, r.Start)
                    pg[m.group(1) + m.group(2)] = c.Information(3)
                    ys[m.group(1) + m.group(2)] = c.Information(6)
        finally:
            d.Close(0)
    finally:
        w.Quit()
    res = {"kept": kept(pg), "y": ys}
    (OUT / ("word_fine.json" if "--fine" in sys.argv else "word.json")).write_text(json.dumps(res, indent=0), encoding="utf-8")
    show(res["kept"])


def oxi(exe):
    sys.path.insert(0, str(REPO / "tools" / "metrics"))
    os.environ["OXI_GDI_EXE"] = os.path.abspath(exe)
    import measure_pagination_oxi as MO
    MO.RENDERER = os.path.abspath(exe)
    r = MO.measure_doc(str(OUT / ("gridbottom_fine.docx" if "--fine" in sys.argv else "gridbottom.docx")))
    pg, ys = {}, {}
    for p, els in r["pages"].items():
        for e in els:
            m = re.search(r"([ST])(\d+)$", e["text"].strip())
            if m and m.group(1) + m.group(2) not in pg:
                pg[m.group(1) + m.group(2)] = int(p)
                ys[m.group(1) + m.group(2)] = e["y"]
    res = {"kept": kept(pg), "y": ys}
    (OUT / ("oxi_fine.json" if "--fine" in sys.argv else "oxi.json")).write_text(json.dumps(res, indent=0), encoding="utf-8")
    show(res["kept"])


def show(k):
    print("".join("K" if k[h] else "." for h in HS), f"(H {HS[0]}..{HS[-1]})")
    last = max([h for h in HS if k[h]], default=None)
    print("largest kept H:", last)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    cmd = sys.argv[1]
    if cmd == "gen": gen()
    elif cmd == "word": word()
    else: oxi(sys.argv[2])
