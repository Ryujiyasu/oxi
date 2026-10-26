# -*- coding: utf-8 -*-
"""A pct-wide table whose tblGrid sums to less: grid or width?

blind-I en/reference__00f0e7f3 (tabloid landscape calendar, compat 12): both
tables say tblW 5000 pct (100%), tblLayout fixed, and a tblGrid / tcW in dxa
that sums to 5000 tw = 250pt (the generator wrote the pct numbers as twips).
Word draws the tables across the full 1080pt text width with the columns in
grid proportion (300 / 300 / 4400 -> 64.8 / 64.8 / 950.4); Oxi kept the 250pt
grid, so the day-name column was one glyph wide and the page ran to three.
Faithful slices; both tables changed alike:
    asis      fixed, pct 5000
    autofit   tblLayout removed
    pct2500   tblW 2500 pct (50%)
    dxa       tblW 5000 dxa (the grid's own sum)
    autodxa   tblLayout removed and tblW 5000 dxa
    c15       asis with compatibilityMode 15
    tcwpct    asis with every tcW as the same share in pct (300 dxa -> 300 pct = 6%)
    tcwnone   asis with the tcW removed
    gridnotype  asis with the (non-schema) w:type attribute dropped from every gridCol
    near / near2  grid and tcW x4.3 (1075pt) / x4.0 (1000pt)
Read: page count and the main table's horizontal extent (the widest stroke)
and its first inner vertical rule x (the day / name boundary).

    python tools/metrics/_pb_tbl_pct_grid_gen.py   (PB_OXI=<exe>[;<exe>] adds Oxi)
"""
import glob, json, os, re, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
SRC = Path(glob.glob(str(REPO / "pipeline_data/docx_corpus/en/reference/00f0e7f3*.docx"))[0])
OUT = REPO / "pipeline_data" / "_pb_tbl_pct_grid"
ARMS = ["asis", "autofit", "pct2500", "dxa", "autodxa", "c15", "tcwpct", "tcwnone", "gridnotype", "near", "near2"]


def variant(name):
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    d = parts["word/document.xml"].decode("utf-8")
    if name in ("autofit", "autodxa"):
        d = d.replace('<w:tblLayout w:type="fixed"/>', "")
    if name == "pct2500":
        d = d.replace('<w:tblW w:w="5000" w:type="pct"/>', '<w:tblW w:w="2500" w:type="pct"/>')
    if name in ("dxa", "autodxa"):
        d = d.replace('<w:tblW w:w="5000" w:type="pct"/>', '<w:tblW w:w="5000" w:type="dxa"/>')
    if name == "tcwpct":
        d = re.sub(r'<w:tcW w:w="(\d+)" w:type="dxa"/>', r'<w:tcW w:w="\1" w:type="pct"/>', d)
    if name == "gridnotype":
        d = d.replace('<w:gridCol w:w="300" w:type="dxa"/>', '<w:gridCol w:w="300"/>').replace(
            '<w:gridCol w:w="4400" w:type="dxa"/>', '<w:gridCol w:w="4400"/>').replace(
            '<w:gridCol w:w="5000" w:type="dxa"/>', '<w:gridCol w:w="5000"/>')
    if name in ("near", "near2"):
        # the grid and tcW x4.3 (1075pt, 5pt short of the share) / x4.0 (1000pt)
        k = 4.3 if name == "near" else 4.0
        d = re.sub(r'(<w:(?:gridCol|tcW) w:w=")(\d+)(")', lambda m: f"{m.group(1)}{round(int(m.group(2)) * k)}{m.group(3)}", d)
    if name == "tcwnone":
        d = re.sub(r'<w:tcW [^>]*/>', "", d)
    parts["word/document.xml"] = d.encode("utf-8")
    if name == "c15":
        s = parts["word/settings.xml"].decode("utf-8")
        s = re.sub(r'(w:name="compatibilityMode" w:uri="[^"]*" w:val=")12"', r'\g<1>15"', s)
        parts["word/settings.xml"] = s.encode("utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return p


def read_word(pdf):
    import pymupdf
    doc = pymupdf.open(pdf)
    pg = doc[0]
    hs = [d["rect"] for d in pg.get_drawings() if d["rect"].height < 2 and d["rect"].width > 30]
    vs = sorted({round(d["rect"].x0, 1) for d in pg.get_drawings() if d["rect"].width < 2 and d["rect"].height > 8})
    x0 = min(r.x0 for r in hs); x1 = max(r.x1 for r in hs)
    return f"pages {len(doc)} table {x0:.1f}..{x1:.1f} verticals {vs[:4]}"


def read_oxi(exe, docx):
    t = tempfile.mkdtemp(); lay = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "q"), "--dump-layout=" + lay], capture_output=True)
    pages = json.load(open(lay, encoding="utf-8"))["pages"]
    b = [e for e in pages[0]["elements"] if e.get("type") == "border"]
    hs = [e for e in b if e.get("h", 0) < 2 and e.get("w", 0) > 30]
    vs = sorted({round(e["x"], 1) for e in b if e.get("w", 0) < 2 and e.get("h", 0) > 8})
    if not hs:
        return f"pages {len(pages)} no borders"
    return f"pages {len(pages)} table {min(e['x'] for e in hs):.1f}..{max(e['x'] + e['w'] for e in hs):.1f} verticals {vs[:4]}"


if __name__ == "__main__":
    exes = [e for e in os.environ.get("PB_OXI", "").split(";") if e]
    for name in os.environ.get("PB_ARMS", ",".join(ARMS)).split(","):
        p = variant(name); pdf = p.with_suffix(".pdf")
        if not pdf.exists():
            import win32com.client
            w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
            try:
                d = w.Documents.Open(str(p), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
            finally:
                w.Quit()
        print(f"{name:8} W {read_word(pdf)}" + "".join(f"   O {read_oxi(e, p)}" for e in exes), flush=True)
