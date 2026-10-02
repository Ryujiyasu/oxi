# -*- coding: utf-8 -*-
"""S1631 probe: where does Word put a BODY ruby line's base and annotation
when the line is not on a typed grid?

Host package: JA educational__09422f63 (its fonts/styles; the section's docGrid
is dropped and every paragraph says snapToGrid 0).  Body per arm: a plain MS
Mincho 10.5 line «前の行», the ruby line (base B pt «漢字»+ruby «かんじ» hps H,
hpsRaise R), a plain line «後の行».
Read (Word PDF / Oxi glyph dump): baselines of «前», the base «漢», the ruby
«か» and «後»; printed as offsets from «前»'s baseline.

    python tools/metrics/_pb_bodyruby_gen.py <renderer.exe> [ENV=1 ...]
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/educational/09422f63e991d48f.docx"
OUT = REPO / "tests/fixtures/bodyruby" / ("plain" if os.environ.get("BR_PLAIN") else "ruby")
# (base half-points, ruby hps half-points, hpsRaise half-points)
ARMS = [(21, 10, 20), (21, 10, 30), (28, 14, 36), (60, 30, 58), (60, 30, 80)]
if os.environ.get("BR_SIZES"):
    # BR_SIZES=21,28,40,60,80: base sizes in half-points (ruby hps/raise kept proportional)
    ARMS = [(b, b // 2, b) for b in map(int, os.environ["BR_SIZES"].split(","))]
# BR_FONT=HG丸ｺﾞｼｯｸM-PRO : the face (default MS Mincho); BR_ARMS=32:16:30,... : (base,hps,raise) half-points
_F = {"hg": "HG丸ｺﾞｼｯｸM-PRO", "mincho": "ＭＳ 明朝", "gothic": "ＭＳ ゴシック"}.get(os.environ.get("BR_FONT", "mincho"), "ＭＳ 明朝")
FONT = f'<w:rFonts w:ascii="{_F}" w:eastAsia="{_F}" w:hAnsi="{_F}"/>'
if os.environ.get("BR_ARMS"):
    ARMS = [tuple(int(v) for v in a.split(":")) for a in os.environ["BR_ARMS"].split(",")]
# BR_GRID=1: keep the host's docGrid (lines 360) and let the paragraphs snap to it
PPR = ('<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' if os.environ.get("BR_GRID")
       else '<w:pPr><w:snapToGrid w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>')


def name(a):
    return f"b{a[0]}_h{a[1]}_r{a[2]}"


def plain(t):
    return f'<w:p>{PPR}<w:r><w:rPr>{FONT}<w:sz w:val="21"/></w:rPr><w:t>{t}</w:t></w:r></w:p>'


def ruby(b, h, r):
    rp = f'<w:rPr>{FONT}<w:sz w:val="{b}"/></w:rPr>'
    if os.environ.get("BR_PLAIN"):
        # BR_PLAIN=1: the same line with no ruby at all (isolates the base placement)
        return f'<w:p>{PPR}<w:r>{rp}<w:t>漢字本文</w:t></w:r></w:p>'
    raise_xml = f'<w:hpsRaise w:val="{r}"/>' if r else ''  # r=0: omit hpsRaise (Word's default raise)
    align = os.environ.get("BR_ALIGN", "distributeSpace")  # BR_ALIGN=distributeLetter|center|left|right
    base_t = os.environ.get("BR_BASE", "漢字")  # BR_BASE=漢 BR_RT=かんじ: a ruby wider than its base
    rt_t = os.environ.get("BR_RT", "かんじ")
    return (f'<w:p>{PPR}<w:r>{rp}<w:ruby><w:rubyPr><w:rubyAlign w:val="{align}"/><w:hps w:val="{h}"/>'
            f'{raise_xml}<w:hpsBaseText w:val="{b}"/><w:lid w:val="ja-JP"/></w:rubyPr>'
            f'<w:rt><w:r><w:rPr>{FONT}<w:sz w:val="{h}"/></w:rPr><w:t>{rt_t}</w:t></w:r></w:rt>'
            f'<w:rubyBase><w:r>{rp}<w:t>{base_t}</w:t></w:r></w:rubyBase></w:ruby></w:r>'
            f'<w:r>{rp}<w:t>本文</w:t></w:r></w:p>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    b0 = doc.index("<w:body>") + len("<w:body>")
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    if not os.environ.get("BR_GRID"):
        sect = re.sub(r"<w:docGrid[^>]*/>", "", sect)
    for a in ARMS:
        body = plain("前の行") + ruby(*a) + plain("後の行")
        xml = doc[:b0] + body + sect + "</w:body></w:document>"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"{name(a)}.docx").write_bytes(buf.getvalue())


def word_baselines():
    import fitz, win32com.client
    res = {}
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    try:
        for a in ARMS:
            tmp = os.path.join(tempfile.mkdtemp(), name(a) + ".docx")
            shutil.copy(OUT / f"{name(a)}.docx", tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            pdf = tmp[:-5] + ".pdf"
            d.SaveAs2(pdf, 17)
            d.Close(0)
            got = {}
            for b in fitz.open(pdf)[0].get_text("rawdict")["blocks"]:
                for l in b.get("lines", []):
                    for s in l["spans"]:
                        for c in s["chars"]:
                            if c["c"] in "前漢か後" and c["c"] not in got:
                                got[c["c"]] = c["origin"][1]
                            if os.environ.get("BR_X") and c["c"] in "前漢字かんじ本" and ("x" + c["c"]) not in got:
                                got["x" + c["c"]] = c["origin"][0]
            res[name(a)] = got
    finally:
        w.Quit()
    return res


def oxi_baselines(exe, envs):
    env = dict(os.environ)
    env.update(kv.split("=", 1) for kv in envs)
    res = {}
    for a in ARMS:
        t = tempfile.mkdtemp()
        out = os.path.join(t, "g.json")
        subprocess.run([exe, str(OUT / f"{name(a)}.docx"), os.path.join(t, "p"), "--dump-glyphs=" + out],
                       capture_output=True, env=env)
        got = {}
        for g in json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]:
            if g["char"] in "前漢か後" and g["char"] not in got:
                got[g["char"]] = g["top"] + 0.859 * g["font_size"]
            if os.environ.get("BR_X") and g["char"] in "前漢字かんじ本" and ("x" + g["char"]) not in got:
                got["x" + g["char"]] = g["x"]
        res[name(a)] = got
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    gen()
    W = word_baselines()
    O = oxi_baselines(os.path.abspath(sys.argv[1]), sys.argv[2:])
    for a in ARMS:
        n = name(a); wv, ov = W[n], O[n]
        rel = lambda d: ({k: round(d[k] - d["前"], 2) for k in "漢か後" if k in d and "前" in d}
                         | ({k: round(d[k] - d["x漢"], 2) for k in ("x字", "xか", "xん", "xじ") if k in d and "x漢" in d} if os.environ.get("BR_X") == "1" else {})
                         # BR_X=2: x of every read glyph from the plain line's left edge («前»)
                         | ({k: round(d[k] - d["x前"], 2) for k in ("x漢", "x字", "xか", "xん", "xじ", "x本") if k in d and "x前" in d} if os.environ.get("BR_X") == "2" else {}))
        print(f"{n:16} Word {rel(wv)}")
        print(f"{'':16} Oxi  {rel(ov)}")
