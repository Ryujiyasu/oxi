# -*- coding: utf-8 -*-
"""Where does Word wrap ONE real justified line?  (reports__0013bcb8, S1629 follow-up)

Faithful slice: the host's own package (fonts, styles, compat 15) and the runs
of the paragraph that holds «adipiscing elit. Donec lorem orci, mattis sit amet
semper vel,», cut so that line starts the paragraph (no first-line indent).
The page's right margin is binary-searched (1tw) for the smallest text width at
which Word still puts «vel,» on line 1, and the same width is asked of Oxi.

    python tools/metrics/_pb_jline_slice_gen.py <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / os.environ.get("JL_SRC", "pipeline_data/docx_corpus/en/reports/0013bcb8b619da89.docx")
PARA = os.environ.get("JL_PARA", "(Walsh, 2007). Lorem ipsum dolor sit amet, consectetur")
CUT = os.environ.get("JL_CUT", "adipiscing")
START = "adipiscing elit. Donec lorem orci"
LAST = os.environ.get("JL_LAST", "vel,")  # JL_LAST=vel. swaps the comma for a period
OUT = REPO / "tests/fixtures/jline_slice" / os.environ.get("JL_TAG", "default")


def slice_doc(text_w_tw):
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    para = next(m.group(0) for m in re.finditer(r"<w:p[ >].*?</w:p>", doc, re.S)
                if PARA in re.sub(r"<[^>]+>", "", m.group(0)))
    # drop every run before the one holding "adipiscing", and the first-line indent
    runs = list(re.finditer(r"<w:r[ >].*?</w:r>", para, re.S))
    k = next(i for i, r in enumerate(runs) if CUT in re.sub(r"<[^>]+>", "", r.group(0)))
    # cut INSIDE the run so the line starts exactly at CUT
    first = runs[k].group(0)
    tpos = first.index(CUT)
    tstart = first.rfind(">", 0, tpos) + 1
    first = first[:tstart] + first[tpos:]
    ppr = re.findall(r"<w:pPr>.*?</w:pPr>", para, re.S)[0]
    ppr = re.sub(r'<w:ind [^>]*/>', "", ppr)
    body = "<w:p>" + ppr + first + "".join(r.group(0) for r in runs[k + 1:]) + "</w:p>"
    if os.environ.get("JL_SWAP"):
        a, b = os.environ["JL_SWAP"].split("=>")
        body = body.replace(a, b, 1)
    if os.environ.get("JL_NOSCALE"):
        # JL_NOSCALE=1: drop the w:w character scale (105) from every run
        body = re.sub(r'<w:w w:val="\d+"/>', "", body)
    b0 = doc.index("<w:body>") + len("<w:body>")
    left = 1134
    page_w = left * 2 + text_w_tw
    sect = (f'<w:sectPr><w:pgSz w:w="{page_w}" w:h="16838"/><w:pgMar w:top="1134" w:right="{left}" '
            f'w:bottom="1134" w:left="{left}" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    xml = doc[:b0] + body + sect + "</w:body></w:document>"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(copy.copy(item), data)
    return buf.getvalue()


def first_line_word(pdf):
    import fitz
    pg = fitz.open(pdf)[0]
    lines = sorted(((l["spans"][0]["origin"][1], "".join(s["text"] for s in l["spans"]))
                    for b in pg.get_text("dict")["blocks"] for l in b.get("lines", []) if l["spans"]))
    return lines[0][1]


def first_line_oxi(exe, docx):
    t = tempfile.mkdtemp()
    out = os.path.join(t, "l.json")
    subprocess.run([exe, docx, os.path.join(t, "p"), "--dump-layout=" + out], capture_output=True)
    p = json.load(open(out, encoding="utf-8"))["pages"][0]
    rows = {}
    for e in p["elements"]:
        if e.get("type") == "text":
            rows.setdefault(round(e["y"], 1), []).append(e)
    y0 = min(rows)
    return "".join(e["text"] for e in sorted(rows[y0], key=lambda e: e["x"]))


def main(exe):
    import win32com.client
    OUT.mkdir(parents=True, exist_ok=True)
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    cache = {}

    def word_keeps(tw):
        if tw not in cache:
            p = OUT / f"w{tw}.docx"
            p.write_bytes(slice_doc(tw))
            tmp = os.path.join(tempfile.mkdtemp(), p.name)
            shutil.copy(p, tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            pdf = tmp.replace(".docx", ".pdf")
            d.SaveAs2(pdf, 17)
            d.Close(0)
            cache[tw] = LAST in first_line_word(pdf).split()
        return cache[tw]

    try:
        lo, hi = (3800, 4800) if os.environ.get("JL_NOSCALE") else (4000, 5000)
        if os.environ.get("JL_BOUNDS"):
            lo, hi = map(int, os.environ["JL_BOUNDS"].split(","))
        assert not word_keeps(lo) and word_keeps(hi), (word_keeps(lo), word_keeps(hi))
        while hi - lo > 1:
            mid = (lo + hi) // 2
            if word_keeps(mid):
                hi = mid
            else:
                lo = mid
        print(f"Word keeps «{LAST}» from {hi} tw = {hi / 20:.2f}pt (wraps at {lo / 20:.2f}pt)")
        for tw in (hi - 40, hi - 1, hi, 4408):
            p = OUT / f"w{tw}.docx"
            p.write_bytes(slice_doc(tw))
            print(f"  {tw / 20:.2f}pt: Oxi line 1 ends «{first_line_oxi(exe, str(p)).rstrip()[-24:]}»  Word keeps={word_keeps(tw)}")
    finally:
        w.Quit()


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(os.path.abspath(sys.argv[1]))
