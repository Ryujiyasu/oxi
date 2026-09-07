# -*- coding: utf-8 -*-
"""Where is the line edge for a run with w:spacing (tracking) on a character grid?

Three document lines disagree: -4 on 167853 p3 fits 21 characters against the true
edge 235.65 (a 20-cell floor would refuse), +2 on 167853 p2 wraps a 37th character
that the true edge would take, -8 on 0ea3ec86 p9 sits on the boundary. The faithful
slice (0ea3ec86 package, 2-column charSpace 2048, cell 11.5) sweeps spacing x length.

Original kinsoku doc-string follows.

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
OUT = os.path.join(REPO, "pipeline_data", "_pb_trackedge")
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


KANA = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわ"

_ROT = [0]

def kana(n):
    # each call rotates the kana so every arm starts with its own 3 characters
    _ROT[0] += 3
    r = _ROT[0] % len(KANA)
    return (KANA[r:] + KANA[:r]) * 3

# (label, spacing twips, text)   cell = 11.5 + 2*spacing/20
ARMS = []
for sp, cell in ((-8, 10.7), (-4, 11.1), (-2, 11.3), (0, 11.5), (2, 11.7), (4, 11.9)):
    nw = int(235.65 // cell)          # chars that fit against the true edge
    nf = int(230.0 // cell)           # chars that fit against the 20-cell section floor
    for n in sorted({nf, nf + 1, nw, nw + 1}):
        ARMS.append(("s%+d_n%d" % (sp, n), sp, kana(n)[:n] + "字"))            # n kana + 字: does the (n+1)th fit? no marks
    ARMS.append(("s%+d_m%d" % (sp, nw), sp, (lambda k: k[:nw - 2] + "、" + k[nw - 2:nw] + "字")(kana(nw))))  # one 、 : half-cell pull-in against the edge
    ARMS.append(("s%+d_sp%d" % (sp, nw), sp, "SPACE" + kana(nw - 1)[:nw - 1] + "字"))   # untracked leading 　 + tracked text
ARMS.append(("SECTBREAK", 0, ""))
for sp, cell in ((-8, 10.7), (-4, 11.1), (0, 11.5), (2, 11.7)):
    nw = int(493.2 // cell)
    nf = int(483.0 // cell)
    for n in sorted({nf, nf + 1, nw, nw + 1}):
        ARMS.append(("c1_s%+d_n%d" % (sp, n), sp, kana(n)[:n] + "字"))
    ARMS.append(("c1_s%+d_m%d" % (sp, nw), sp, (lambda k: k[:nw - 2] + "、" + k[nw - 2:nw] + "字")(kana(nw))))
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
    sect1 = re.sub(r'w:rsid\w*="[^"]*" ?', "", re.sub(r"<w:(header|footer)Reference[^>]*/>", "", secs[5]))  # continuous, 1 column, charSpace 2048
    for label, sp, text in ARMS:
        if label == "SECTBREAK":
            body += "<w:p><w:pPr>" + sect + "</w:pPr></w:p>"
            sect = sect1
            continue
        rpr = '<w:rPr><w:spacing w:val="%d"/></w:rPr>' % sp if sp else ""
        if text.startswith("SPACE"):
            body += ('<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">　</w:t></w:r>'
                     '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (rpr, text[5:]))
        else:
            body += '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (rpr, text)
        body += "<w:p/>"
    new = head + body + sect + "</w:body></w:document>"
    out = os.path.join(OUT, "trackedge.docx")
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
    src = os.path.join(OUT, "trackedge.docx")
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
    lines = [l for l in lines if l[4].strip()]
    for label, sp, text in ARMS:
        if label == "SECTBREAK":
            continue
        full = text.replace("SPACE", "　")
        key = full.replace("　", "")[:3]
        for i, (pno, col, y, wdt, t) in enumerate(lines):
            tt = t.replace(" ", "").replace("　", "")
            if tt.startswith(key):
                nxt = lines[i + 1][4].replace(" ", "") if i + 1 < len(lines) else ""
                print("%-11s cell=%.1f line1=%2d w=%6.1f of %d | ...%s / %s" % (label, 11.5 + 2 * sp / 20.0, len(tt), wdt, len(full.replace("　", "")), tt[-5:], nxt[:5]))
                break


if __name__ == "__main__":
    {"gen": gen, "pdf": pdf}[sys.argv[1]]()
