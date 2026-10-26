# -*- coding: utf-8 -*-
"""Page-number fields in headers / footers: Word's digit face and size against Oxi's.

Documents whose PAGE / NUMPAGES field carries `\\* MERGEFORMAT` and whose cached
result run is formatted differently from the field code (the case
OXI_FIELD_MERGEFORMAT changes). Page 1, digits in the top or bottom 15% of the
page: Word's (PDF from a fresh Word process, cached in
pipeline_data/_field_result_census/) against each exe's glyph dump.

    python tools/metrics/_field_result_census.py <exe>[;<exe>] <doc> [<doc> ...]
        (<doc> = corpus path without .docx, e.g. en/forms/00447890dc66a9d8)
"""
import json, os, subprocess, sys, tempfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "_field_result_census"


def zone(y, h):
    return y < 0.15 * h or y > 0.85 * h


def word(docx, pdf):
    if not pdf.exists():
        import win32com.client
        w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(str(docx), ReadOnly=True); d.ExportAsFixedFormat(str(pdf), 17); d.Close(False)
        finally:
            w.Quit()
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    h = pg.rect.height
    return sorted({(round(c["origin"][1], 1), s["font"].split("-")[0].replace("PSMT", "").replace("MT", ""), round(s["size"], 1))
                   for b in pg.get_text("rawdict")["blocks"] for l in b.get("lines", []) for s in l["spans"]
                   for c in s["chars"] if c["c"].isdigit() and zone(c["origin"][1], h)})


def oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True, timeout=1800)
    pg = json.load(open(g, encoding="utf-8"))["pages"][0]
    return sorted({(round(c["baseline"], 1), c["font_family"], round(c["font_size"], 1))
                   for c in pg["glyphs"] if c["char"].isdigit() and zone(c["baseline"], pg["height"])})


if __name__ == "__main__":
    exes = [e for e in sys.argv[1].split(";") if e]
    OUT.mkdir(parents=True, exist_ok=True)
    for doc in sys.argv[2:]:
        docx = REPO / "pipeline_data/docx_corpus" / f"{doc}.docx"
        print(f"== {doc}")
        print(f"   W {word(docx, OUT / (doc.replace('/', '_') + '.pdf'))}")
        for e in exes:
            print(f"   O {oxi(e, docx)}", flush=True)
