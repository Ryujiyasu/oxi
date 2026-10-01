# -*- coding: utf-8 -*-
"""S1620 probe: where does a CJK table's border width go without a typed grid?

Host package: JA forms__000af17f (its styles/settings/fonts, docGrid with no
type).  Body: two MS Mincho 10.5 lines, a bordered table, two lines after it.
Arms (name = grid_sz_rows):
  grid  nt = docGrid linePitch with no type (the host's), none = no docGrid
  sz    table border w:sz (4 = 0.5pt, 12 = 1.5pt)
  rows  1 or 2
Read: Word PDF (rule centres, text baselines) vs the Oxi dump (rule top edges +
width, text line tops) -- printed as Word - Oxi per item, so a uniform offset is
the font's ascent and anything that changes across the table is the border.

    python tools/metrics/_pb_notype_tbl_gen.py gen
    python tools/metrics/_pb_notype_tbl_gen.py cmp <renderer.exe> [ENV=1 ...]
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/forms/000af17f67209a4f.docx"
OUT = REPO / "tests/fixtures/notype_tbl"
ARMS = [(g, sz, r) for g in ("nt", "none") for sz in (4, 12) for r in (1, 2)]
FONT = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr>'


def name(a):
    return f"{a[0]}_{a[1]}_{a[2]}"


def para(t):
    return f"<w:p><w:r>{FONT}<w:t>{t}</w:t></w:r></w:p>"


def table(sz, rows):
    b = "".join(f'<w:{s} w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>'
                for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    tr = "".join(f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{para(f"表{i}左")}</w:tc>'
                 f'<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{para(f"表{i}右")}</w:tc></w:tr>'
                 for i in range(rows))
    return (f'<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblBorders>{b}</w:tblBorders>'
            f'<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            f'{tr}</w:tbl>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    b0 = doc.index("<w:body>") + len("<w:body>")
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    for a in ARMS:
        s = sect if a[0] == "nt" else re.sub(r"<w:docGrid[^>]*/>", "", sect)
        body = (para("前行一あいうえお") + para("前行二かきくけこ") + table(a[1], a[2])
                + para("後行一さしすせそ") + para("後行二たちつてと"))
        xml = doc[:b0] + body + s + "</w:body></w:document>"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"{name(a)}.docx").write_bytes(buf.getvalue())
    print("ok", len(ARMS))


def word_pdf():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    try:
        for a in ARMS:
            pdf = OUT / f"{name(a)}.pdf"
            if pdf.exists():
                continue
            tmp = os.path.join(tempfile.mkdtemp(), name(a) + ".docx")
            shutil.copy(OUT / f"{name(a)}.docx", tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            d.SaveAs2(str(pdf), 17)
            d.Close(0)
    finally:
        w.Quit()


def cmp(exe, envs):
    import fitz
    word_pdf()
    env = dict(os.environ)
    env.update(kv.split("=", 1) for kv in envs)
    for a in ARMS:
        pg = fitz.open(str(OUT / f"{name(a)}.pdf"))[0]
        wr = sorted({round((dr["rect"].y0 + dr["rect"].y1) / 2, 2) for dr in pg.get_drawings()
                     if dr["rect"].height < 2.5 and dr["rect"].width > 100})
        wt = {}
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).strip()
                if t:
                    wt[t[:3]] = l["spans"][0]["origin"][1]
        dump = OUT / "_o.json"
        subprocess.run([os.path.abspath(exe), str(OUT / f"{name(a)}.docx"), str(OUT / "_o"), "--dump-layout=" + str(dump)],
                       capture_output=True, env=env)
        p = json.load(open(dump, encoding="utf-8"))["pages"][0]
        bw = a[1] / 8.0
        orr = sorted({round(e["y"] + bw / 2, 2) for e in p["elements"]
                      if e.get("type") == "border" and e.get("h", 0) < 0.01 and e.get("w", 0) > 100})
        by_y = {}
        for e in p["elements"]:
            if e.get("type") == "text" and e["text"].strip():
                by_y.setdefault((round(e["y"], 2), round(e["x"] // 200)), []).append(e)
        ot = {}
        for (y, _), es in by_y.items():
            ot.setdefault("".join(e["text"] for e in sorted(es, key=lambda e: e["x"])).strip()[:3], y)
        texts = " ".join(f"{k}:{wt[k] - ot[k]:+.2f}" for k in wt if k in ot)
        rules = " ".join(f"{x - y:+.2f}" for x, y in zip(wr, orr))
        print(f"{name(a):12} rules W-O [{rules}] ({len(wr)}/{len(orr)})  text W-O {texts}")
        dump.unlink()
    for x in OUT.glob("_o*.png"):
        x.unlink()


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if sys.argv[1] == "gen":
        gen()
    else:
        cmp(sys.argv[2], sys.argv[3:])
