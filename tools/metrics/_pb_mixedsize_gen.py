# -*- coding: utf-8 -*-
"""Mixed font sizes in ONE CJK line: where does Word put the baseline of a
larger run, and how much does it grow the line?  (09422f63 「危険なのはいつ？」:
a 14pt '？' in a 12pt line -- Oxi aligns glyph TOPS, so the big glyph's
baseline is 1.6pt low and the line 1.2pt too tall.)

Host 09422f63 (HG丸ｺﾞｼｯｸM-PRO), docGrid dropped, snapToGrid 0. Body per arm:
«前の行» (base pt), the test line «いつ» + BIG «？» (big pt) + «です», «後の行».
Arms: MS_ARMS=12:14,12:16,12:18,12:12,10.5:14 (base:big). MS_RUBY=1 puts ruby
on «いつ» as in the real line.
Read: Word PDF glyph origins (baselines) and bbox tops of «い» and «？», and the
baseline of «後»; Oxi the same from --dump-glyphs (top; baseline = top + ascent is
NOT known, so tops are compared to Word's tops).

    python tools/metrics/_pb_mixedsize_gen.py <renderer.exe>
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/educational/09422f63e991d48f.docx"
OUT = REPO / "tests/fixtures/mixedsize"
ARMS = [tuple(float(v) for v in a.split(":")) for a in os.environ.get("MS_ARMS", "12:14,12:16,12:18,12:12,10.5:14").split(",")]
RUBY = bool(os.environ.get("MS_RUBY"))
F = "HG丸ｺﾞｼｯｸM-PRO"
FONT = f'<w:rFonts w:ascii="{F}" w:eastAsia="{F}" w:hAnsi="{F}"/>'
PPR = '<w:pPr><w:snapToGrid w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'


def name(a):
    return f"b{a[0]:g}_big{a[1]:g}{'_ruby' if RUBY else ''}{'_sym' if SYM else ''}"


SYM = bool(os.environ.get("MS_SYM"))  # MS_SYM=1: the BIG run is «◆» in Segoe UI Symbol (a Latin-font symbol in a CJK line)
SYMFONT = '<w:rFonts w:ascii="Segoe UI Symbol" w:eastAsia="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol" w:cs="Segoe UI Symbol"/>'


def run(t, pt, sym=False):
    f = SYMFONT if sym else FONT
    return f'<w:r><w:rPr>{f}<w:sz w:val="{int(pt * 2)}"/></w:rPr><w:t>{t}</w:t></w:r>'


def plain(t, pt):
    return f'<w:p>{PPR}{run(t, pt)}</w:p>'


def test_line(base, big):
    if RUBY:
        rp = f'<w:rPr>{FONT}<w:sz w:val="{int(base * 2)}"/></w:rPr>'
        head = (f'<w:r>{rp}<w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="{int(base)}"/>'
                f'<w:hpsRaise w:val="{int(base * 2 - 2)}"/><w:hpsBaseText w:val="{int(base * 2)}"/><w:lid w:val="ja-JP"/></w:rubyPr>'
                f'<w:rt><w:r><w:rPr>{FONT}<w:sz w:val="{int(base)}"/></w:rPr><w:t>いつ</w:t></w:r></w:rt>'
                f'<w:rubyBase><w:r>{rp}<w:t>何時</w:t></w:r></w:rubyBase></w:ruby></w:r>')
    else:
        head = run("いつ", base)
    big_run = run("◆", big, True) if SYM else run("？", big)
    return f'<w:p>{PPR}{head}{big_run}{run("です", base)}</w:p>'


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    b0 = doc.index("<w:body>") + len("<w:body>")
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    sect = re.sub(r"<w:docGrid[^>]*/>", "", sect)
    for a in ARMS:
        body = plain("前の行", a[0]) + test_line(*a) + plain("後の行", a[0])
        xml = doc[:b0] + body + sect + "</w:body></w:document>"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"{name(a)}.docx").write_bytes(buf.getvalue())


def word():
    import fitz, win32com.client
    res = {}
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        for a in ARMS:
            n = name(a); tmp = os.path.join(tempfile.mkdtemp(), n + ".docx"); shutil.copy(OUT / f"{n}.docx", tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False); pdf = tmp[:-5] + ".pdf"; d.SaveAs2(pdf, 17); d.Close(0)
            got = {}
            for b in fitz.open(pdf)[0].get_text("rawdict")["blocks"]:
                for l in b.get("lines", []):
                    for s in l["spans"]:
                        for c in s["chars"]:
                            if c["c"] in "前い？◆で後" and c["c"] not in got:
                                got[c["c"]] = (round(c["bbox"][1], 2), round(c["origin"][1], 2))
            res[n] = got
    finally:
        w.Quit()
    return res


def oxi(exe):
    res = {}
    for a in ARMS:
        n = name(a); t = tempfile.mkdtemp(); out = os.path.join(t, "g.json")
        subprocess.run([exe, str(OUT / f"{n}.docx"), os.path.join(t, "p"), "110", "--dump-glyphs=" + out], capture_output=True)
        got = {}
        for g in json.load(open(out, encoding="utf-8"))["pages"][0]["glyphs"]:
            if g["char"] in "前い？◆で後" and g["char"] not in got:
                got[g["char"]] = round(g["top"], 2)
        res[n] = got
    return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    gen()
    W = word()
    O = oxi(os.path.abspath(sys.argv[1])) if len(sys.argv) > 1 else {}
    print("Word: (glyph top, baseline) relative to «前»'s baseline;  Oxi: glyph top relative to «前»'s top")
    for a in ARMS:
        n = name(a); wv = W[n]; p = wv["前"][1]
        print(f"{n:16} Word " + "  ".join(f"{k}:top{wv[k][0]-p:+.2f}/base{wv[k][1]-p:+.2f}" for k in "い？◆で後" if k in wv))
        if n in O:
            ov = O[n]; pt = ov["前"]
            print(f"{'':16} Oxi  " + "  ".join(f"{k}:top{ov[k]-pt:+.2f}" for k in "い？◆で後" if k in ov) + f"   (Word tops rel 前-top: " + "  ".join(f"{k}:{wv[k][0]-wv['前'][0]:+.2f}" for k in "い？◆で後" if k in wv) + ")")
