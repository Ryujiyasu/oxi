# -*- coding: utf-8 -*-
"""Which run's formatting does Word give a complex field's displayed result?

en/policies__009ad56e footer1: `of NUMPAGES` -- the begin / instrText / separate
runs carry no rPr (Calibri via docDefaults), the cached result «6» and the end
run are Times New Roman, `\\* MERGEFORMAT`. Word prints «Page 1 of 6» all in
Times; Oxi drew the count in Calibri and sized the footer line by Calibri
(13.43 vs 12.65), so the footer and the notes above it sat 0.8 high.
Arms (NUMPAGES only):
    asis           as is
    nomerge        the `\\* MERGEFORMAT` switch removed
    charformat     `\\* CHARFORMAT` instead
    resarial       the cached result run in Arial (code runs still bare)
    resarial_nm    resarial without MERGEFORMAT
    codearial      the instrText run in Arial, result Times
Read: page 1's footer line -- the face of the last digit, the line's baseline.

    python tools/metrics/_pb_field_result_face_gen.py   (PB_OXI=<exe>[;<exe>] adds Oxi)
"""
import json, os, re, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/policies/009ad56ee7b95f4e.docx"
OUT = REPO / "pipeline_data" / "_pb_field_result_face"
ARMS = ["asis", "nomerge", "charformat", "resarial", "resarial_nm", "codearial"]
TNR = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
ARIAL = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'


def variant(name):
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    for n in ("word/footer1.xml", "word/footer2.xml"):
        x = parts[n].decode("utf-8")
        i = x.find("NUMPAGES")
        assert i > 0
        k = max(x.rfind("<w:r>", 0, i), x.rfind("<w:r ", 0, i))   # the instrText run
        sep = x.find('w:fldCharType="separate"', i)
        res = re.compile(r"<w:r[ >]").search(x, sep).start()      # the cached result run
        res_end = x.find("</w:r>", res)
        head, instr, mid, result, tail = x[:k], x[k:sep], x[sep:res], x[res:res_end], x[res_end:]
        if name in ("nomerge", "resarial_nm"):
            instr = instr.replace("\\* MERGEFORMAT", "")
        if name == "charformat":
            instr = instr.replace("\\* MERGEFORMAT", "\\* CHARFORMAT")
        if name in ("resarial", "resarial_nm"):
            assert TNR in result
            result = result.replace(TNR, ARIAL)
        if name == "codearial":
            instr = re.sub(r"(<w:r[^>]*>)<w:instrText", lambda m: m.group(1) + f"<w:rPr>{ARIAL}</w:rPr><w:instrText", instr, count=1)
        parts[n] = (head + instr + mid + result + tail).encode("utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return p


def read_word(pdf):
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    chars = [(c["origin"][1], c["origin"][0], c["c"], s["font"]) for b in pg.get_text("rawdict")["blocks"]
             for l in b.get("lines", []) for s in l["spans"] for c in s["chars"] if c["origin"][1] > 745 and c["c"].strip()]
    last = max(chars, key=lambda c: c[1])
    return f"base {last[0]:.2f} last «{last[2]}» {last[3]}"


def read_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    chars = [c for c in json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"] if c["baseline"] > 745 and c["char"].strip()]
    last = max(chars, key=lambda c: c["x"])
    return f"base {last['baseline']:.2f} last «{last['char']}» {last['font_family']}"


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
        print(f"{name:12} W {read_word(pdf)}" + "".join(f"   O {read_oxi(e, p)}" for e in exes), flush=True)
