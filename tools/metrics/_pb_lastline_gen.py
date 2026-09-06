# -*- coding: utf-8 -*-
"""How far does Word SQUEEZE the LAST line of a paragraph (no filler follows)?

0ea3ec86 p6 「技術を習得した手話奉仕員の養成・研修を行う。」: 22 characters on a
20-cell column, 19 gaps shrunk 0.4pt each, ・ 0.67em, 。 half; p34 「ーションや
在宅就労を促進する。講習期間２年。」: 22 characters ALL at 10.56 (x0.917).
Sweep the paragraph-final overflow in cells, with and without marks.

Original doc-string of the kinsoku probe follows.

When does Word pull a unit 「X、」 / 「X。」 onto a full line of a
two-column compat-14 compressPunctuation document (reference__0ea3ec86)?

The Word PDF refuses 「お、」 after 「とする人に特別障害者手当（国）がある。な」
(three marks to compress, 2.0 cells against a 1.51 demand) and 「事、」 after
「食事等の介護、調理、洗濯及び掃除などの家」, yet keeps 「た。」 on
「るための法律（障害者総合支援法）」とされた。」 with three brackets at half.
A faithful slice (the document's own styles / settings / fonts, one 2-column
section with charSpace 2048) carries synthetic 20-cell lines followed by the
unit, with the number, kind and position of the marks swept.

    python _pb_kinsokufinal_gen.py gen
    python _pb_kinsokufinal_gen.py pdf
"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
SRC = os.path.join(REPO, "pipeline_data", "docx_corpus", "ja", "reference", "0ea3ec86480140c2.docx")
OUT = os.path.join(REPO, "pipeline_data", "_pb_lastline")
sys.stdout.reconfigure(encoding="utf-8")

KANA = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわ"
FILL = "以下は次の行に流れる本文であって行末の判定には関わらない文字列を続ける。"


def line20(marks, kind="、"):
    """20 characters: kana with `marks` marks at spread positions (never first/last)."""
    chars = list(KANA[:20])
    pos = {1: [10], 2: [6, 13], 3: [5, 10, 15], 4: [4, 8, 12, 16]}.get(marks, [])
    for i, p in enumerate(pos):
        chars[p] = kind[i % len(kind)]
    return "".join(chars)


ARMS = [
    ("k0_x", line20(0) + "X"),                 # 21 chars, no mark, paragraph end
    ("k0_xx", line20(0) + "XX"),               # 22
    ("k0_xxx", line20(0) + "XXX"),             # 23
    ("k0_x。", line20(0) + "X。"),             # 22 with final 。
    ("k0_xx。", line20(0) + "XX。"),           # 23
    ("k0_xxx。", line20(0) + "XXX。"),         # 24
    ("k1_x。", line20(1) + "X。"),
    ("k1_xx。", line20(1) + "XX。"),
    ("k2_x。", line20(2) + "X。"),
    ("k2_xx。", line20(2) + "XX。"),
    ("k2_xxx。", line20(2) + "XXX。"),
    ("k1_x", line20(1) + "X"),
    ("k2_xx", line20(2) + "XX"),
    ("p6", "技術を習得した手話奉仕員の養成・研修を行う。"),
    ("p34", "ーションや在宅就労を促進する。講習期間２年。"),
    ("k0_x_more", line20(0) + "X" + "以下続く"),     # control: NOT a paragraph end (filler)
]
ARMS = [(l, t.replace("X", "字")) for l, t in ARMS]


def gen():
    os.makedirs(OUT, exist_ok=True)
    z = zipfile.ZipFile(SRC)
    doc = z.read("word/document.xml").decode("utf-8")
    secs = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)
    sect = secs[4]  # 2 columns, charSpace 2048
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)
    sect = re.sub(r'w:rsid\w*="[^"]*" ?', "", sect)
    head = doc[:doc.index("<w:body>") + len("<w:body>")]
    body = ""
    for label, text in ARMS:
        body += '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text
        body += "<w:p/>"
    new = head + body + sect + "</w:body></w:document>"
    out = os.path.join(OUT, "lastline.docx")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as o:
        for item in z.infolist():
            if item.filename == "word/document.xml":
                o.writestr(item, new.encode("utf-8"))
            elif item.filename.startswith("word/header") or item.filename.startswith("word/footer"):
                continue
            else:
                data = z.read(item.filename)
                if item.filename == "word/_rels/document.xml.rels":
                    data = re.sub(rb'<Relationship [^>]*Target="(header|footer)\d*\.xml"[^>]*/>', b"", data)
                if item.filename == "[Content_Types].xml":
                    data = re.sub(rb'<Override [^>]*PartName="/word/(header|footer)\d*\.xml"[^>]*/>', b"", data)
                o.writestr(item, data)
    print("wrote", out)


def pdf():
    import fitz
    import win32com.client as w
    src = os.path.join(OUT, "lastline.docx")
    out = src[:-5] + ".pdf"
    app = w.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(src, ReadOnly=True)
        d.ExportAsFixedFormat(out, 17)
        d.Close(False)
    finally:
        app.Quit()
    doc = fitz.open(out)
    lines = []
    for pno in range(len(doc)):
        page = doc[pno]
        mid = page.rect.width / 2
        for b in page.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l["spans"]).replace(" ", "")
                if t:
                    lines.append((pno, 0 if l["bbox"][0] < mid else 1, round(l["bbox"][1], 1), round(l["bbox"][2] - l["bbox"][0], 1), t))
    lines.sort()
    firsts = {t[:12]: (label, t) for label, t in ARMS}
    for i, (pno, col, y, wdt, t) in enumerate(lines):
        for k, (label, full) in firsts.items():
            if t.startswith(k[:8]):
                nxt = lines[i + 1][4] if i + 1 < len(lines) else ""
                print("%-10s line1=%2d chars w=%6.1f %s || line2: %s" % (label, len(t), wdt, t, nxt[:6]))


if __name__ == "__main__":
    {"gen": gen, "pdf": pdf}[sys.argv[1]]()
