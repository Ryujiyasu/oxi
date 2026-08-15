# -*- coding: utf-8 -*-
"""Does the paragraph MARK's font grow a line that already has text?

technical__002c1ffa's TOC entries resolve like this: the runs carry only
`<w:noProof/>` — no font, no size — so they inherit docDefaults' Times New Roman
at the style's 9pt, and Word draws them as TimesNewRomanPSMT 9.00. The paragraph
MARK, though, carries `asciiTheme="minorHAnsi"` (Calibri) at sz=22 (11pt), and
Word emits a Calibri 11.04 space span at the end of every such line.

Word's line pitch there is 10.1 ≈ Times New Roman 9's natural 9.97 — the 11pt
mark does NOT lift it. Oxi's is 10.5. Half a point per line is what decides
whether a two-line entry clears the page bottom, and that one decision cascades
into 641 mis-paged paragraphs across the document's 368 pages.

Each arm is one page holding a paragraph long enough to wrap once, so the line
height is the y difference between its two lines — measured directly rather than
inferred from where a page happens to break.

  python _pb_markline_gen.py gen
  python _pb_markline_gen.py pdf      # Word truth
  python _pb_markline_gen.py oxi      # Oxi, same arms
"""
import json
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_markline")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS  # noqa: E402

# (name, mark font (None = none declared), mark size half-points, run font, run size)
ARMS = [
    ("a_spec",       "Calibri", 22, None, None),   # the specimen: bare runs, 11pt mark
    ("b_mark_same",  "Calibri", 18, None, None),   # mark at the text's own 9pt
    ("c_no_mark",    None,      0,  None, None),   # no mark rPr at all
    ("d_mark_16",    "Calibri", 32, None, None),   # a much larger mark
    ("e_runs_named", "Calibri", 22, "Times New Roman", 18),  # runs name the font
    ("f_mark_tnr",   "Times New Roman", 22, None, None),     # 11pt mark, same family
    # ★the mark is ignored by BOTH (all arms above land on one pitch), so the
    # specimen's extra 0.15pt is elsewhere. Its mark also names an EAST ASIAN
    # theme font, and its style adds spacing before + a leader tab; these arms
    # add those one at a time.
    ("g_mark_ea",    "Calibri", 22, None, None),   # mark also names an eastAsia font
    ("h_style_sb",   "Calibri", 22, None, None),   # + spacing before=40
    ("i_leader_tab", "Calibri", 22, None, None),   # + right leader-dot tab
]
LINE = ("Determination must be revoked if there is a change to the "
        "responsible person's cost percentage and the registrar is notified "
        "of that change within the period allowed by the regulations")


def docx():
    return os.path.join(OUT, "markline.docx")


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (name, mfont, msz, rfont, rsz) in enumerate(ARMS):
        body.append(
            '<w:p><w:pPr>%s<w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
            '<w:sz w:val="20"/></w:rPr></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>'
            "<w:t>M%02d</w:t></w:r></w:p>"
            % ("<w:pageBreakBefore/>" if ai > 0 else "", ai))
        mark_rpr = ""
        if name == "g_mark_ea":
            mark_rpr = ('<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"'
                        ' w:eastAsia="MS Mincho"/><w:sz w:val="22"/></w:rPr>')
            mfont = None
        elif mfont:
            mark_rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
                        '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
                        % (mfont, mfont, msz, msz))
        run_rpr = "<w:rPr><w:noProof/></w:rPr>"
        if rfont:
            run_rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
                       '<w:sz w:val="%d"/></w:rPr>' % (rfont, rfont, rsz))
        # w:ind narrows the column so the sentence wraps exactly once
        sb = ' w:before="40"' if name in ("h_style_sb", "i_leader_tab") else ""
        tab = ('<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="7088"/></w:tabs>'
               if name == "i_leader_tab" else "")
        body.append(
            '<w:p><w:pPr>%s<w:spacing%s w:after="0" w:line="240"'
            ' w:lineRule="auto"/><w:ind w:left="2098" w:right="567"/>%s</w:pPr>'
            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (tab, sb, mark_rpr, run_rpr, LINE))
        body.append(
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
            ' w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
            '<w:sz w:val="20"/></w:rPr></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>'
            "<w:t>E%02d</w:t></w:r></w:p>" % ai)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11907" w:h="16839"/>'
           '<w:pgMar w:top="1418" w:right="2410" w:bottom="1418" w:left="2410" '
           'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    # docDefaults = Times New Roman with NO size, exactly as the specimen writes
    # it; the paragraph style below supplies the 9pt.
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              "<w:docDefaults><w:rPrDefault><w:rPr>"
              '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman"'
              ' w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
              "</w:rPr></w:rPrDefault>"
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240"'
              ' w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              '<w:name w:val="Normal"/><w:rPr><w:sz w:val="18"/></w:rPr></w:style>'
              "</w:styles>")
    with zipfile.ZipFile(docx(), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/document.xml", doc)
    print("wrote", docx(), len(ARMS), "arms")


def report(per, who):
    print("== %s ==" % who)
    print("%-14s %-16s %-6s %-16s %10s %8s"
          % ("arm", "mark font", "mark", "run font", "line pitch", "lines"))
    for ai, (name, mfont, msz, rfont, rsz) in enumerate(ARMS):
        got = per.get(ai)
        if not got:
            print("%-14s MISSING" % name)
            continue
        ys = got
        pitch = (ys[1] - ys[0]) if len(ys) > 1 else 0.0
        print("%-14s %-16s %-6s %-16s %10.2f %8d"
              % (name, mfont or "(none)", msz / 2.0 if msz else "-",
                 rfont or "(inherit)", pitch, len(ys)))


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
    per = {}
    for ai in range(len(ARMS)):
        if ai >= doc.page_count:
            break
        ys, seen = [], set()
        for bl in doc[ai].get_text("dict")["blocks"]:
            if bl["type"] != 0:
                continue
            for ln in bl["lines"]:
                t = "".join(s["text"] for s in ln["spans"]).strip()
                if not t or t.startswith(("M", "E")):
                    continue
                y = round(ln["bbox"][1], 2)
                if y not in seen:
                    seen.add(y)
                    ys.append((y, [(s["font"], round(s["size"], 2)) for s in ln["spans"]]))
        ys.sort()
        if ys:
            per[ai] = [y for y, _ in ys]
            print("   %-14s spans: %s" % (ARMS[ai][0], ys[0][1][:3]))
    report(per, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "markline_oxi.json")
    subprocess.run([GDI, docx(), os.path.join(tempfile.gettempdir(), "ml"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    pages = json.load(open(out, encoding="utf-8"))["pages"]
    per = {}
    for ai in range(len(ARMS)):
        if ai >= len(pages):
            break
        ys = sorted({round(e["y"], 2) for e in pages[ai]["elements"]
                     if e.get("type") == "text"
                     and (e.get("text") or "").strip()
                     and not (e.get("text") or "").strip().startswith(("M", "E"))})
        if ys:
            per[ai] = ys
    report(per, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "pdf": pdf}[sys.argv[1]]()
