# -*- coding: utf-8 -*-
"""S1624 probe: does a ruby line on a typed grid grow by the ruby expansion,
or only to the next whole grid cell?

Host package: JA forms__01c5a769 (docGrid lines 360 = 18pt, no compat element,
TableGrid style).  One document per arm; each holds a 1-column table of six
rows, every row one paragraph «氏名» at base size B (half-points) carrying a
ruby «ふりがな» of size 16 with hpsRaise R, then the same six paragraphs in the
body.  Arms: B in (21, 28), R in (20, 36, 52, 68, 84, 100) -- one row / body
paragraph per R, so one document per B.
Read: Word PDF row rules (row height per R) and Info(6) steps of the body
paragraphs; Oxi dump rules and line tops.

    python tools/metrics/_pb_gridruby_gen.py gen
    python tools/metrics/_pb_gridruby_gen.py cmp <renderer.exe> [ENV=1 ...]
"""
import copy, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/forms/01c5a769623fcec5.docx"
OUT = REPO / "tests/fixtures/gridruby"
BASES = (21, 28)
RAISES = (20, 36, 52, 68, 84, 100)


def ruby_para(base, raise_):
    r = f'<w:rPr><w:sz w:val="{base}"/></w:rPr>'
    return (f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>{r}<w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/>'
            f'<w:hps w:val="16"/><w:hpsRaise w:val="{raise_}"/><w:hpsBaseText w:val="{base}"/><w:lid w:val="ja-JP"/></w:rubyPr>'
            f'<w:rt><w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝" w:hint="eastAsia"/>'
            f'<w:sz w:val="16"/></w:rPr><w:t>ふりがな</w:t></w:r></w:rt>'
            f'<w:rubyBase><w:r>{r}<w:t>氏名{raise_}</w:t></w:r></w:rubyBase></w:ruby></w:r></w:p>')


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    b0 = doc.index("<w:body>") + len("<w:body>")
    sect = re.findall(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    for b in BASES:
        rows = "".join(f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{ruby_para(b, r)}</w:tc></w:tr>'
                       for r in RAISES)
        tbl = ('<w:tbl><w:tblPr><w:tblStyle w:val="a3"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/></w:tblPr>'
               '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' + rows + '</w:tbl>')
        body = tbl + "<w:p/>" + "".join(ruby_para(b, r) for r in RAISES) + "<w:p/>"
        xml = doc[:b0] + body + sect + "</w:body></w:document>"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(copy.copy(item), data)
        (OUT / f"b{b}.docx").write_bytes(buf.getvalue())
    print("ok", len(BASES))


def word():
    import win32com.client, fitz
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    res = {}
    try:
        for b in BASES:
            tmp = os.path.join(tempfile.mkdtemp(), f"b{b}.docx")
            shutil.copy(OUT / f"b{b}.docx", tmp)
            d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
            ys = []
            for i in range(1, d.Paragraphs.Count + 1):
                rg = d.Paragraphs(i).Range
                if rg.Information(12):  # wdWithInTable
                    continue
                # a ruby paragraph's Range.Text is its EQ field code, so take every
                # body paragraph and keep the ones between the two empty ones
                ys.append(d.Range(rg.Start, rg.Start).Information(6))
            ys = ys[1:1 + len(RAISES)]
            pdf = OUT / f"b{b}.pdf"
            d.SaveAs2(str(pdf), 17)
            d.Close(0)
            pg = fitz.open(str(pdf))[0]
            rules = sorted({round((r["rect"].y0 + r["rect"].y1) / 2, 2) for r in pg.get_drawings()
                            if r["rect"].height < 2.5 and r["rect"].width > 100})
            res[b] = {"rows": [round(y - x, 2) for x, y in zip(rules, rules[1:])],
                      "body": [round(y - x, 2) for x, y in zip(ys, ys[1:])]}
    finally:
        w.Quit()
    (OUT / "word.json").write_text(json.dumps(res), encoding="utf-8")
    return res


def oxi(exe, envs):
    env = dict(os.environ)
    env.update(kv.split("=", 1) for kv in envs)
    res = {}
    for b in BASES:
        dump = OUT / "_o.json"
        subprocess.run([os.path.abspath(exe), str(OUT / f"b{b}.docx"), str(OUT / "_o"), "--dump-layout=" + str(dump)],
                       capture_output=True, env=env)
        p = json.load(open(dump, encoding="utf-8"))["pages"][0]
        rules = sorted({round(e["y"], 2) for e in p["elements"]
                        if e.get("type") == "border" and e.get("h", 0) < 0.01 and e.get("w", 0) > 100})
        tops = sorted({round(e["y"], 2) for e in p["elements"] if e.get("type") == "text" and "氏" in e["text"]})
        body = [y for y in tops if y > rules[-1]] if rules else tops
        res[b] = {"rows": [round(y - x, 2) for x, y in zip(rules, rules[1:])],
                  "body": [round(y - x, 2) for x, y in zip(body, body[1:])]}
        dump.unlink()
    for x in OUT.glob("_o*.png"):
        x.unlink()
    return res


def cmp(exe, envs):
    wf = OUT / "word.json"
    wr = json.loads(wf.read_text(encoding="utf-8")) if wf.exists() else word()
    orr = oxi(exe, envs)
    for b in BASES:
        print(f"base {b / 2}pt  raises {[r / 2 for r in RAISES]}")
        print(f"   rows  W {wr[str(b)]['rows'] if str(b) in wr else wr[b]['rows']}")
        print(f"         O {orr[b]['rows']}")
        print(f"   body  W {wr[str(b)]['body'] if str(b) in wr else wr[b]['body']}")
        print(f"         O {orr[b]['body']}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if sys.argv[1] == "gen":
        gen()
    else:
        cmp(sys.argv[2], sys.argv[3:])
