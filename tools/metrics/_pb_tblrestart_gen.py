# -*- coding: utf-8 -*-
"""S1621 probe: where is a continued table's top rule drawn on the next page?

Host package: EN reports__0013bcb8 (its TabloKlavuzu = Table Grid style, docGrid
with no type).  Body: an exact spacer of H pt, then a 30-row Calibri table whose
borders come from the table STYLE only (no tblBorders), so it splits onto p2.
Arms: H in SPACERS (moves the split row).  Read on p2: the first rule (Word PDF
centre vs Oxi drawn edge + half width) and the first row's text (Word baseline
- Oxi line top), so a rule drawn too low shows as a rule delta that the text
delta does not share.

    python tools/metrics/_pb_tblrestart_gen.py gen
    python tools/metrics/_pb_tblrestart_gen.py cmp <renderer.exe> [ENV=1 ...]
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/reports/0013bcb8b619da89.docx"
OUT = REPO / "tests/fixtures/tblrestart"
SPACERS = [450]
ARMS = [(h, cb, tt) for h in SPACERS for cb in ("none", "bottom", "top") for tt in ("style", "none")]
FONT = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/></w:rPr>'


def para(t):
    return f"<w:p><w:r>{FONT}<w:t>{t}</w:t></w:r></w:p>"


def table(rows, cb="bottom", tt="style"):
    # as reports__0013bcb8's tables: tblBorders turns the side and inside edges
    # off, so the top rule comes from the style (tt="none" turns it off too);
    # cb = the cell's own border (none / bottom / top)
    tc = "" if cb == "none" else f'<w:tcBorders><w:{cb} w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>'
    tr = "".join(f'<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>{tc}</w:tcPr>{para(f"Row{i}A")}</w:tc>'
                 f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>{tc}</w:tcPr>{para(f"Row{i}B")}</w:tc></w:tr>'
                 for i in range(rows))
    off = "".join(f'<w:{e} w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                  for e in (("top",) if tt == "none" else ()) + ("left", "bottom", "right", "insideH", "insideV"))
    return ('<w:tbl><w:tblPr><w:tblStyle w:val="TabloKlavuzu"/><w:tblW w:w="6000" w:type="dxa"/>'
            f'<w:tblBorders>{off}</w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' + tr + '</w:tbl>')


def name(a):
    return f"h{a[0]}_{a[1]}_{a[2]}"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    b0 = doc.index("<w:body>") + len("<w:body>")
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[0]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    for a in ARMS:
        sp = (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{a[0] * 20}" w:lineRule="exact"/></w:pPr>'
              f'<w:r>{FONT}<w:t>Spacer</w:t></w:r></w:p>')
        xml = doc[:b0] + sp + table(30, a[1], a[2]) + para("After") + sect + "</w:body></w:document>"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"{name(a)}.docx").write_bytes(buf.getvalue())
    print("ok", len(ARMS))


def rows_on(page_pdf, page_oxi):
    """Word baseline - Oxi line top for every 'Row..' line on the page."""
    wt = {}
    for b in page_pdf.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if t.startswith("Row"):
                wt[t] = l["spans"][0]["origin"][1]
    ot = {}
    for e in page_oxi["elements"]:
        if e.get("type") == "text" and e["text"].startswith("Row"):
            ot.setdefault(e["text"], e["y"])
    return [(k, round(wt[k] - ot[k], 2)) for k in sorted(wt, key=wt.get) if k in ot]


def cmp(exe, envs):
    import fitz, win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    try:
        for a in ARMS:
            pdf = OUT / f"{name(a)}.pdf"
            if not pdf.exists():
                tmp = os.path.join(tempfile.mkdtemp(), f"{name(a)}.docx")
                shutil.copy(OUT / f"{name(a)}.docx", tmp)
                d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
                d.SaveAs2(str(pdf), 17)
                d.Close(0)
    finally:
        w.Quit()
    env = dict(os.environ)
    env.update(kv.split("=", 1) for kv in envs)
    for a in ARMS:
        doc = fitz.open(str(OUT / f"{name(a)}.pdf"))
        dump = OUT / "_o.json"
        subprocess.run([os.path.abspath(exe), str(OUT / f"{name(a)}.docx"), str(OUT / "_o"), "--dump-layout=" + str(dump)],
                       capture_output=True, env=env)
        pages = json.load(open(dump, encoding="utf-8"))["pages"]
        p1 = rows_on(doc[0], pages[0])
        p2 = rows_on(doc[1], pages[1])
        ref = sorted(d for _, d in p1)[len(p1) // 2] if p1 else None
        wr = sorted({round((dr["rect"].y0 + dr["rect"].y1) / 2, 2) for dr in doc[1].get_drawings()
                     if dr["rect"].height < 2.5 and dr["rect"].width > 100})[:2]
        print(f"{name(a):18} p1 ref {ref}  p2 first {p2[0] if p2 else None}  "
              f"shift {round(p2[0][1] - ref, 2) if p2 and ref is not None else None}  W p2 rules {wr}")
        dump.unlink()
    for x in OUT.glob("_o*.png"):
        x.unlink()


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if sys.argv[1] == "gen":
        gen()
    else:
        cmp(sys.argv[2], sys.argv[3:])
