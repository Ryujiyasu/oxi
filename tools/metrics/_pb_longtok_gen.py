# -*- coding: utf-8 -*-
"""Where does Word put a Latin token too long to fit the rest of the line?

This is the HEAD of tokyoshugyo's divergence chain (`_kojin_rowgeom.py scan`
puts the first departure on p4, and every later page inherits some of it). The
paragraph there ends with an e-gov URL:

  Word  …詳細はこちら                                    <- token not started
        （https://…/Procedure?CLASSNAME=GTAMSTDETAIL&id  <- full line
        =4950000009642&…）をご確認ください。
  Oxi   …詳細はこちら（https:                            <- token started here
        //shinsei.e-gov.go.jp/search/servlet/Procedure?  <- a quarter empty
        CLASSNAME=…&fromGTAEGOVMSTDETAIL=true）をご確認くださ
        い。

so Oxi spends one line more. Two questions, swept separately:

  (1) does the token START on the partly-filled line, or move down whole?
  (2) once it owns a full line, where does the break fall -- at the margin
      (character level) or at a punctuation opportunity inside the token?

Arm = one CJK prefix length (which sets the room left on the first line) x one
token flavour (no punctuation / slashes every 10). Each arm is its own page.

    python _pb_longtok_gen.py gen
    python _pb_longtok_gen.py pdf      # Word truth
    python _pb_longtok_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_longtok")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

FACE = "ＭＳ 明朝"
# tokyoshugyo's URL run carries w:hAnsi="ＭＳ 明朝", so the Latin glyphs come from
# the Mincho face itself. Using a separate Latin face here (Century was the first
# try) drags an unrelated metrics question into the answer: Word fit 102 of those
# chars per line and Oxi 120+, which is an advance-width gap, not a wrap rule.
ASCII_FACE = "ＭＳ 明朝"
SZ_HP = 21                 # 10.5pt, as in tokyoshugyo
PITCH = 360                # twips = 18pt

# 60-char tokens: longer than a line at 10.5pt (line box 510.2pt / ~5.25pt per
# half-width char = ~97 chars, so use 120 to be safely over one full line).
PLAIN = ("abcdefghij" * 12)
SLASHY = "/".join(["abcdefghi"] * 12) + "/abcdefghi"
TOKENS = [("plain", PLAIN), ("slashy", SLASHY)]
# CJK prefix lengths: 10.5pt CJK is 10.5pt per char, so each step eats 10.5pt of
# the first line. 0 puts the token at the line start (question 2 only); the rest
# leave progressively less room (question 1).
PREFIXES = [0, 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44]
ARMS = [(t, n) for t, _ in TOKENS for n in PREFIXES]


def docx():
    return os.path.join(OUT, "longtok.docx")


def para(text, ppr=""):
    return ('<w:p><w:pPr>%s</w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (ppr, ASCII_FACE, ASCII_FACE, FACE, SZ_HP, SZ_HP, text))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (tname, n) in enumerate(ARMS):
        tok = dict(TOKENS)[tname]
        body.append(para("A%02dZ" % ai,
                         "<w:pageBreakBefore/>" if ai else ""))
        body.append(para("あ" * n + tok + "。おわり"))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/>'
           '<w:pgMar w:top="1985" w:right="851" w:bottom="1701" w:left="851" '
           'w:header="851" w:footer="992" w:gutter="0"/>'
           '<w:docGrid w:type="lines" w:linePitch="%d" w:charSpace="0"/>'
           "</w:sectPr></w:body></w:document>" % PITCH)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/>'
              "</w:rPr></w:rPrDefault></w:docDefaults>"
              '<w:style w:type="paragraph" w:default="1" w:styleId="a">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="%d"/></w:rPr></w:style>'
              "</w:styles>" % (ASCII_FACE, FACE, ASCII_FACE, SZ_HP))
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms")


def report(per, who):
    print("== %s ==" % who)
    print("%-4s %-7s %-4s %-7s %-6s %s"
          % ("arm", "token", "pfx", "on_pfx", "lines", "owner line ends"))
    for ai, (tname, n) in enumerate(ARMS):
        lines = per.get(ai)
        if not lines:
            print("%-4d %-7s %-4d MISSING" % (ai, tname, n))
            continue
        tok = dict(TOKENS)[tname]
        # token chars sitting on the PREFIX line: 0 means Word moved the whole
        # token down. Counted from the text, not guessed from a suffix test.
        on_pfx = 0
        for ln in lines:
            if "あ" in ln:
                tail = ln[ln.rindex("あ") + 1:]
                on_pfx = len(tail)
                break
        owner = next((l for l in lines if tok[:6] in l), None)
        cut = ""
        if owner:
            taken = owner[owner.index(tok[:6]):]
            cut = "%d chars (…%s)" % (len(taken), taken[-6:])
        print("%-4d %-7s %-4d %-7d %-6d %s"
              % (ai, tname, n, on_pfx, len(lines), cut))


def _collect(pagelines):
    """pagelines: list of (page, [line texts]) -> arm -> lines of its paragraph."""
    per = {}
    for lines in pagelines:
        ai = None
        for t in lines:
            s = t.strip()
            if s.startswith("A") and s.endswith("Z") and s[1:-1].isdigit():
                ai = int(s[1:-1])
                per[ai] = []
                continue
            if ai is not None and s:
                per[ai].append(s)
    return per


def pdf():
    import fitz
    import win32com.client as w
    out = docx().replace(".docx", ".pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(docx(), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(out, 17)
    finally:
        d.Close(False)
        app.Quit()
    doc = fitz.open(out)
    pages = []
    for pi in range(doc.page_count):
        lines = []
        for b in doc[pi].get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                lines.append((round(ln["bbox"][1], 2),
                              "".join(s["text"] for s in ln["spans"])))
        pages.append([t for _, t in sorted(lines)])
    report(_collect(pages), "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "longtok_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "lt"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = []
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        rows = {}
        for e in pg["elements"]:
            if e["type"] == "text":
                rows.setdefault(round(e["y"], 1), []).append((e["x"], e.get("text") or ""))
        pages.append(["".join(t for _, t in sorted(v)) for _, v in sorted(rows.items())])
    report(_collect(pages), "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    elif sys.argv[1] == "pdf":
        pdf()
    else:
        gen()
