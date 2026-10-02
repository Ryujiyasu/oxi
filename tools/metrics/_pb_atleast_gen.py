# -*- coding: utf-8 -*-
"""Probe (2026-10-02): where is the first baseline of a Latin line whose
w:spacing lineRule="atLeast" exceeds the natural line (no typed grid)?

Host: en/educational 0061215a (margin.top 36, docGrid linePitch 360 no type,
no header). Body per arm: one paragraph, Arial bold S pt, spacing
line=L atLeast, text «Line one AgQ», then a plain 11pt paragraph «after».
Read (Word PDF span origin / Oxi --dump-glyphs top + winAscent): the first
baseline from the top margin, and the gap to «after».

    python tools/metrics/_pb_atleast_gen.py <renderer.exe>
    AL_ARMS=27:330,27:400,24:240,24:500,21:330   (sz half-points : line twips)
    AL_RULE=atLeast|exact|auto
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/en/educational/0061215af6c4e0ab.docx"
OUT = REPO / "tests/fixtures/atleast"
RULE = os.environ.get("AL_RULE", "atLeast")
ARMS = [tuple(int(v) for v in a.split(":")) for a in os.environ.get("AL_ARMS", "27:330,27:400,27:500,24:330,22:330,27:240").split(",")]
FONT = '<w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial"/><w:b/>'


def name(a):
    return f"s{a[0]}_l{a[1]}"


def body(sz, line):
    rule = "" if RULE == "auto" else f' w:lineRule="{RULE}"'
    return (f'<w:p><w:pPr><w:spacing w:after="0" w:line="{line}"{rule}/><w:rPr>{FONT}<w:sz w:val="{sz}"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr>{FONT}<w:sz w:val="{sz}"/></w:rPr><w:t>Line one AgQ</w:t></w:r></w:p>'
            f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>after</w:t></w:r></w:p>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    b0 = doc.index("<w:body>") + len("<w:body>")
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    for a in ARMS:
        xml = doc[:b0] + body(*a) + sect + "</w:body></w:document>"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
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
            got = {}
            for b in fitz.open(pdf)[0].get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    for s in l["spans"]:
                        if s["text"].startswith("Line") and "L" not in got:
                            got["L"] = round(s["origin"][1], 2)
                        if s["text"].startswith("after") and "a" not in got:
                            got["a"] = round(s["origin"][1], 2)
            res[name(a)] = got
    finally:
        w.Quit()
    return res


def oxi_read(exe):
    sys.path.insert(0, str(REPO / "tools/metrics"))
    from line_census import win_ascent
    res = {}
    for a in ARMS:
        t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
        subprocess.run([exe, str(OUT / f"{name(a)}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + out], capture_output=True)
        got = {}
        for g in json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]:
            k = {"L": "L", "a": "a"}.get(g["char"])
            if k and k not in got:
                got[k] = round(g["top"] + win_ascent(g["font_family"]) * g["font_size"], 2)
        res[name(a)] = got
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    gen()
    W = word_read(); O = oxi_read(os.path.abspath(sys.argv[1]))
    for a in ARMS:
        n = name(a); wv, ov = W[n], O[n]
        f = lambda d: f"L-margin {d.get('L', 0) - 36:6.2f}  after-L {d.get('a', 0) - d.get('L', 0):6.2f}"
        print(f"{n:10} Word {f(wv)}   | Oxi {f(ov)}")
