# -*- coding: utf-8 -*-
"""Does a bullet marker's face move a grid line's baseline?

blind-H ja/legal__0fde54a6: two «□» bullet paragraphs (marker 游明朝, text
ＭＳ ゴシック 11pt, lines 360) sit 1.19 lower in Word than in Oxi; the same
document's unbulleted lines agree to 0.06. Faithful slices, the marker's
rFonts changed (numbering.xml, abstractNum of numId 1, level 0):
    asis      游明朝
    msmin     ＭＳ 明朝
    msgoth    ＭＳ ゴシック (the text's own face)
    century   Century
    yugoth    游ゴシック
    arial     Arial
    nogrid    asis, the section's docGrid removed
    yu9 / yu14 / yu18, ms14 / ms18   the marker at that size (游明朝 / ＭＳ 明朝)
Read: the baselines of «上記内容について、同意します» and of the unbulleted line
before the list («・事件・事故等…»), Word and Oxi.

    python tools/metrics/_pb_marker_face_grid_gen.py   (PB_OXI=<exe>[;<exe>] adds Oxi)
"""
import glob, json, os, re, subprocess, sys, tempfile, zipfile
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
REPO = Path(__file__).resolve().parents[2]
SRC = Path(glob.glob(str(REPO / "pipeline_data/docx_corpus/ja/legal/0fde54a6*.docx"))[0])
OUT = REPO / "pipeline_data" / "_pb_marker_face_grid"
ARMS = ["asis", "msmin", "msgoth", "century", "yugoth", "arial", "nogrid",
        "yu9", "yu14", "yu18", "ms14", "ms18"]
FACE = {"msmin": "ＭＳ 明朝", "msgoth": "ＭＳ ゴシック", "century": "Century", "yugoth": "游ゴシック", "arial": "Arial"}
ASIS = '<w:rFonts w:ascii="游明朝" w:eastAsia="游明朝" w:hAnsi="游明朝" w:cstheme="minorBidi" w:hint="eastAsia"/>'


def variant(name):
    z = zipfile.ZipFile(SRC)
    parts = {n: z.read(n) for n in z.namelist()}
    num = parts["word/numbering.xml"].decode("utf-8")
    if name[:2] in ("yu", "ms") and name[2:].isdigit():
        f = "游明朝" if name.startswith("yu") else "ＭＳ 明朝"
        v = int(name[2:]) * 2
        num = num.replace(ASIS, f'<w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}" w:hint="eastAsia"/><w:sz w:val="{v}"/><w:szCs w:val="{v}"/>', 1)
    if name in FACE:
        f = FACE[name]
        assert ASIS in num
        num = num.replace(ASIS, f'<w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}" w:cs="{f}" w:hint="eastAsia"/>', 1)
    parts["word/numbering.xml"] = num.encode("utf-8")
    if name == "nogrid":
        d = parts["word/document.xml"].decode("utf-8")
        parts["word/document.xml"] = re.sub(r'<w:docGrid w:type="lines" w:linePitch="360"/>', "", d).encode("utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"{name}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as w:
        for n, b in parts.items():
            w.writestr(n, b)
    return p


KEYS = ("事件・事故等", "同意します")


def read_word(pdf):
    import pymupdf
    pg = pymupdf.open(pdf)[0]
    out = {}
    for b in pg.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"])
            for k in KEYS:
                if k in t and k not in out:
                    out[k] = max(s["origin"][1] for s in l["spans"] if k[0] in s["text"] or k in s["text"])
    return out


def read_oxi(exe, docx):
    t = tempfile.mkdtemp(); g = os.path.join(t, "g.json")
    subprocess.run([exe, str(docx), os.path.join(t, "p"), "--dump-glyphs=" + g], capture_output=True)
    gl = json.load(open(g, encoding="utf-8"))["pages"][0]["glyphs"]
    text = "".join(c["char"] for c in gl)
    out = {}
    for k in KEYS:
        i = text.find(k)
        if i >= 0:
            out[k] = gl[i]["baseline"]
    return out


def fmt(d):
    return "  ".join(f"{k[:4]} {d[k]:.2f}" if k in d else f"{k[:4]} -" for k in KEYS) + (
        f"  (bullet - plain {d[KEYS[1]] - d[KEYS[0]]:.2f})" if all(k in d for k in KEYS) else "")


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
        print(f"{name:8} W {fmt(read_word(pdf))}" + "".join(f"   O {fmt(read_oxi(e, p))}" for e in exes), flush=True)
