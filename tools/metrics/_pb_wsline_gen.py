# -*- coding: utf-8 -*-
"""S1595 probe: the height of a line that holds only whitespace.

Faithful slice of blind-G creative__6fd5a307 (its theme and package kept, no
styles.xml -- Word paints its CJK in MS Mincho): each arm is
  marker para "あ" (18pt) / TEST para / marker para "あ" (18pt)
and the TEST line height is Info(6)(marker2) - Info(6)(TEST).
Arms (run 18pt Hiragino ascii/hAnsi like the source unless noted):
  ideo      U+3000
  ascii2    two U+0020
  ideo_a    U+3000 + あ           (control: a real CJK line)
  ideo_ea   U+3000 with eastAsia ＭＳ 明朝 explicit
  ideo_mark U+3000, paragraph mark rPr sz 36 too
  empty     empty run (sz 36)
  ideo3     three U+3000

    python tools/metrics/_pb_wsline_gen.py gen
    python tools/metrics/_pb_wsline_gen.py measure
"""
import copy, io, json, os, re, shutil, sys, tempfile, zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data/docx_corpus/ja/creative/6fd5a3073c03a7b9.docx"
OUT = REPO / "tests/fixtures/wsline"
RPR = ('<w:rPr><w:rFonts w:ascii="Hiragino Mincho Pro" w:hAnsi="Hiragino Mincho Pro" w:cs="Hiragino Mincho Pro"{ea}/>'
       '<w:sz w:val="36"/><w:szCs w:val="36"/><w:spacing w:val="118"/></w:rPr>')


def para(text, ea="", mark=False):
    rpr = RPR.format(ea=ea)
    ppr = '<w:pPr><w:ind w:left="60"/>' + (rpr if mark else "") + "</w:pPr>"
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


ARMS = {
    "ideo": para("　"),
    "ascii2": para("  "),
    "ideo_a": para("　あ"),
    "ideo_ea": para("　", ea=' w:eastAsia="ＭＳ 明朝"'),
    "ideo_mark": para("　", mark=True),
    "empty": para(""),
    "ideo3": para("　　　"),
}


def gen():
    OUT.mkdir(parents=True, exist_ok=True)
    zin = zipfile.ZipFile(SRC)
    doc = zin.read("word/document.xml").decode("utf-8")
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
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
    (OUT / "wsline.docx").write_bytes(buf.getvalue())
    print("ok")


def measure():
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    try:
        src = OUT / "wsline.docx"
        tmp = os.path.join(tempfile.mkdtemp(), src.name)
        shutil.copy(src, tmp)
        d = w.Documents.Open(tmp, ReadOnly=True, AddToRecentFiles=False)
        try:
            ys = []
            for i in range(1, d.Paragraphs.Count + 1):
                r = d.Paragraphs(i).Range
                c = d.Range(r.Start, r.Start)
                ys.append((c.Information(3), c.Information(6), r.Text[:10]))
            res = {}
            for k, name in enumerate(ARMS):
                a, t, b = ys[3 * k], ys[3 * k + 1], ys[3 * k + 2]
                if a[0] == t[0] == b[0]:
                    res[name] = {"head": round(t[1] - a[1], 2), "test": round(b[1] - t[1], 2)}
                else:
                    res[name] = {"page_split": [a, t, b]}
            print(json.dumps(res, ensure_ascii=False))
            (OUT / "wsline_result.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
        finally:
            d.Close(0)
    finally:
        w.Quit()


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    {"gen": gen, "measure": measure}[sys.argv[1]]()
