# -*- coding: utf-8 -*-
"""Footer census for footers that hold a paragraph after a table.

Lays out every such corpus document with two renderers, keeps the documents
whose footer zone (elements below 80% of the page height) moves between them,
and for those reads Word's footer-zone text baselines from a PDF exported by a
fresh Word process (cached in pipeline_data/_ftr_tbl_census/). Prints, per
changed document, the first page's lowest-band text baselines: Word, A, B.

    python tools/metrics/_ftr_tbl_census.py <exe A> <exe B>
"""
import glob, json, os, re, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_ftr_tbl_census"


def docs():
    for f in sorted(glob.glob(str(REPO / "pipeline_data/docx_corpus/*/*/*.docx"))):
        try:
            z = zipfile.ZipFile(f)
        except Exception:
            continue
        if any(re.match(r"word/footer\d*\.xml$", n) and re.search(r"</w:tbl>\s*<w:p[ >]", z.read(n).decode("utf-8", "replace"))
               for n in z.namelist()):
            yield Path(f)


def oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json"); lay = os.path.join(t, "l.json")
    subprocess.run([exe, str(docx), os.path.join(t, "q"), "--dump-layout=" + lay], capture_output=True, timeout=1800)
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True, timeout=1800)
    pages = json.load(open(lay, encoding="utf-8"))["pages"]
    zone = [sorted((round(e["y"], 2), e.get("type")) for e in p["elements"] if e["y"] > 0.8 * p.get("height", 842)) for p in pages]
    glyphs = json.load(open(g, encoding="utf-8"))["pages"]
    base = [sorted({round(c["baseline"], 2) for c in p["glyphs"] if c["baseline"] > 0.8 * p["height"] and c["char"].strip()}) for p in glyphs]
    return zone, base


def word(docx, pdf):
    if not pdf.exists():
        import win32com.client
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(str(docx), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
        finally:
            w.Quit()
    import pymupdf
    out = []
    for pg in pymupdf.open(pdf):
        h = pg.rect.height
        out.append(sorted({round(c["origin"][1], 2) for b in pg.get_text("rawdict")["blocks"] for l in b.get("lines", [])
                           for s in l["spans"] for c in s["chars"] if c["origin"][1] > 0.8 * h and c["c"].strip()}))
    return out


if __name__ == "__main__":
    a, b = sys.argv[1], sys.argv[2]
    OUT.mkdir(parents=True, exist_ok=True)
    for docx in docs():
        name = f"{docx.parent.parent.name}_{docx.parent.name}_{docx.stem[:8]}"
        za, ba = oxi(a, docx)
        zb, bb = oxi(b, docx)
        if za == zb:
            print(f"{name:30} unchanged", flush=True)
            continue
        w = word(docx, OUT / f"{name}.pdf")
        moved = [i for i in range(min(len(za), len(zb))) if za[i] != zb[i]]
        i = moved[0]
        print(f"{name:30} {len(moved)} pages moved; p{i + 1} W {w[i][-3:] if i < len(w) else '-'}  A {ba[i][-3:]}  B {bb[i][-3:]}", flush=True)
