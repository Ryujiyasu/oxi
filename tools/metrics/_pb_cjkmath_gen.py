# -*- coding: utf-8 -*-
"""S1596 probe: the height of a line holding inline OMML in a CJK document.

Faithful slice of blind-G educational__20d9968b (its styles / settings / theme and
the `lines` docGrid, linePitch 360, kept): each arm is
  marker "あ" / TEST / marker "あ"
and the TEST line height is Info(6)(marker2) - Info(6)(TEST); `head` is the step
into it (the marker line). Arms:
  plain     あいう
  frac      あ + a/b
  fracrad   あ + b / sqrt(a^2+b^2)          (the document's own shape)
  sup       あ + x^2
  fracfrac  あ + (a/b)/(c/d)
  frac_ns   frac with snapToGrid=0
  sub       あ + x_i
    python tools/metrics/_pb_cjkmath_gen.py gen
    python tools/metrics/_pb_cjkmath_gen.py measure
"""
import copy, io, json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/educational/20d9968b2e582be9.docx"
OUT = REPO / "tests/fixtures/cjkmath"
M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'


def mr(t):
    return f'<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>{t}</m:t></m:r>'


def frac(n, d):
    return f"<m:f><m:fPr><m:ctrlPr/></m:fPr><m:num>{n}</m:num><m:den>{d}</m:den></m:f>"


def sup(b, s):
    return f"<m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr><m:e>{b}</m:e><m:sup>{s}</m:sup></m:sSup>"


def sub(b, s):
    return f"<m:sSub><m:sSubPr><m:ctrlPr/></m:sSubPr><m:e>{b}</m:e><m:sub>{s}</m:sub></m:sSub>"


def rad(e):
    return f'<m:rad><m:radPr><m:degHide m:val="on"/><m:ctrlPr/></m:radPr><m:deg/><m:e>{e}</m:e></m:rad>'


def para(text, math=None, nosnap=False):
    ppr = "<w:pPr>" + ('<w:snapToGrid w:val="0"/>' if nosnap else "") + "</w:pPr>"
    body = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
    if math:
        body += f"<m:oMath>{math}</m:oMath>"
    return f"<w:p>{ppr}{body}</w:p>"


ARMS = {
    "plain": para("あいう"),
    "frac": para("あ", frac(mr("a"), mr("b"))),
    "fracrad": para("あ", frac(mr("b"), rad(sup(mr("a"), mr("2")) + mr("+") + sup(mr("b"), mr("2"))))),
    "sup": para("あ", sup(mr("x"), mr("2"))),
    "fracfrac": para("あ", frac(frac(mr("a"), mr("b")), frac(mr("c"), mr("d")))),
    "frac_ns": para("あ", frac(mr("a"), mr("b")), nosnap=True),
    "sub": para("あ", sub(mr("x"), mr("i"))),
}
NUMS = {"a": mr("a"), "b": mr("b"), "a2": sup(mr("a"), mr("2")), "A": mr("A")}
DENS = {"b": mr("b"), "g": mr("g"), "ra": rad(mr("a")), "rab": rad(sup(mr("a"), mr("2")) + mr("+") + sup(mr("b"), mr("2"))),
        "rb": rad(mr("b")), "g2": sup(mr("g"), mr("2"))}
if "--grid" in sys.argv or True:
    for nk, nv in NUMS.items():
        for dk, dv in DENS.items():
            ARMS[f"g_{nk}_{dk}"] = para("あ", frac(nv, dv))


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    if "xmlns:m=" not in head:
        head = head.replace("<w:document ", f"<w:document {M} ", 1)
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    body = ""
    for name, p in ARMS.items():
        body += para("あ" + name) + p + para("あ")
    xml = head + body + sect + "</w:body></w:document>"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(copy.copy(item), data)
    (OUT / "cjkmath.docx").write_bytes(buf.getvalue())
    print("ok")


def measure():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    try:
        src = OUT / "cjkmath.docx"
        tmp = os.path.join(tempfile.mkdtemp(), src.name)
        shutil.copy(src, tmp)
        d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
        try:
            ys = []
            for i in range(1, d.Paragraphs.Count + 1):
                r = d.Paragraphs(i).Range
                c = d.Range(r.Start, r.Start)
                ys.append((c.Information(3), c.Information(6)))
            res = {}
            for k, name in enumerate(ARMS):
                a, t, b = ys[3 * k], ys[3 * k + 1], ys[3 * k + 2]
                res[name] = ({"head": round(t[1] - a[1], 2), "test": round(b[1] - t[1], 2)}
                             if a[0] == t[0] == b[0] else {"page_split": [a, t, b]})
            print(json.dumps(res))
            (OUT / "cjkmath_result.json").write_text(json.dumps(res, indent=1), encoding="utf-8")
            d.ExportAsFixedFormat(str(OUT / "cjkmath.pdf"), 17)
        finally:
            d.Close(0)
    finally:
        w.Quit()


if __name__ == "__main__":
    {"gen": gen, "measure": measure}[sys.argv[1]]()
