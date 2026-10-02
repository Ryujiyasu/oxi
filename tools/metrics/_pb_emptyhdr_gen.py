# -*- coding: utf-8 -*-
"""Probe (2026-10-02): when does an INK-LESS header (one empty paragraph) push
the body top below the top margin?

Host: en/administrative 005be1f9 (titlePg, first header = empty right-aligned
Header paragraph, Calibri 11 -> 13.4pt line; the body's first block is an
inline picture whose top is read straight from the Word PDF).  Arms rewrite
pgMar top / header (twips).  Read: Word's first image top; Oxi's first image y.

    python tools/metrics/_pb_emptyhdr_gen.py <renderer.exe>
    EH_ARMS=567:709,709:709,720:709,851:709,567:851,400:709,567:600
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/administrative/005be1f9a2b6bf33.docx"
OUT = REPO / "tests/fixtures/emptyhdr"
ARMS = [tuple(int(v) for v in a.split(":")) for a in os.environ.get("EH_ARMS", "567:709,709:709,720:709,851:709,567:851,400:709,567:600").split(",")]
# EH_HDR=asis|nojc|center|left|nostyle : rewrite the empty header paragraph's pPr
HDR = os.environ.get("EH_HDR", "asis")


def name(a):
    return f"top{a[0]}_hdr{a[1]}_{HDR}"


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    for a in ARMS:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/header1.xml" and HDR != "asis":
                    x = data.decode("utf-8")
                    if HDR == "nojc":
                        x = x.replace('<w:jc w:val="right"/>', "")
                    elif HDR in ("center", "left"):
                        x = x.replace('<w:jc w:val="right"/>', f'<w:jc w:val="{HDR}"/>')
                    elif HDR == "nostyle":
                        x = re.sub(r"<w:pPr>.*?</w:pPr>", "", x, count=1, flags=re.S)
                    data = x.encode("utf-8")
                if item.filename == "word/document.xml":
                    x = data.decode("utf-8")
                    x = re.sub(r'(<w:pgMar[^>]*w:top=")\d+(")', rf'\g<1>{a[0]}\2', x)
                    x = re.sub(r'(<w:pgMar[^>]*w:header=")\d+(")', rf'\g<1>{a[1]}\2', x)
                    data = x.encode("utf-8")
                zout.writestr(copy.copy(item), data)
        (OUT / f"{name(a)}.docx").write_bytes(buf.getvalue())


def word_read():
    import fitz, win32com.client
    res = {}
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for a in ARMS:
            tmp = os.path.join(tempfile.mkdtemp(), name(a) + ".docx"); shutil.copy(OUT / f"{name(a)}.docx", tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            pdf = tmp[:-5] + ".pdf"; d.SaveAs2(pdf, 17); d.Close(0)
            pg = fitz.open(pdf)[0]
            tops = sorted(round(fitz.Rect(i["bbox"]).y0, 2) for i in pg.get_image_info())
            res[name(a)] = tops[0] if tops else None
    finally:
        w.Quit()
    return res


def oxi_read(exe):
    res = {}
    for a in ARMS:
        t = tempfile.mkdtemp(); out = os.path.join(t, "l.json")
        subprocess.run([exe, str(OUT / f"{name(a)}.docx"), os.path.join(t, "p"), "110", "--dump-layout=" + out], capture_output=True)
        els = json.load(open(out, encoding="utf-8"))["pages"][0]["elements"]
        tops = sorted(round(e["y"], 2) for e in els if e.get("type") == "image")
        res[name(a)] = tops[0] if tops else None
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    gen()
    W = word_read(); O = oxi_read(os.path.abspath(sys.argv[1]))
    for a in ARMS:
        n = name(a)
        print(f"{n:14} margin {a[0]/20:6.2f} header {a[1]/20:6.2f}   Word image top {W[n]}   Oxi {O[n]}")
