# -*- coding: utf-8 -*-
"""Where does Word put the baseline of an EXACT line, in a LATIN document?

S1356 fixed this for CJK: the baseline sits 0.8 x the line height below the
line's top, size-independent, measured with a CJK font in a CJK host. A blind
English document moves 0.7561 -> 0.7803 when that rule is switched off, and
nothing else in 142 switches moves it, so the law is doing the wrong thing
somewhere outside the space it was derived in. Latin is the obvious outside.

    python tools/metrics/_pb_exactline_latin.py [host-stem] [font]

Sweeps the line height against the font size, which is the pair the CJK
derivation found no dependence between. Word exports the PDF; the span
origins are the baselines; the line tops follow from the exact heights.
"""
import os
import re
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
DOCS = REPO / "tools" / "golden-test" / "documents" / "docx"
OUT = REPO / "pipeline_data" / "_pb_exactline_latin.docx"

LINES = [float(x) for x in os.environ.get("EL_LINES", "12,14,16,18,24,30").split(",")]
SIZES = [float(x) for x in os.environ.get("EL_SIZES", "9,10,11,12,14,16,20").split(",")]


def build(host: Path, font: str) -> str:
    z = zipfile.ZipFile(host)
    doc = z.read("word/document.xml").decode("utf-8")
    head = doc[: doc.index("<w:body>") + len("<w:body>")]
    sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)

    paras, plan = [], []
    for line in LINES:
        for size in SIZES:
            tag = f"{int(line)}x{size:g}"
            plan.append((line, size, tag))
            rpr = (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
                   f'<w:sz w:val="{round(size * 2)}"/>'
                   f'<w:szCs w:val="{round(size * 2)}"/></w:rPr>')
            # EL_RULE=auto asks the same questions off the exact rule, which is
            # how you tell an exact-line bug from a line-placement bug.
            rule = os.environ.get("EL_RULE", "exact")
            spacing = (f'w:line="{round(line * 20)}" w:lineRule="exact"'
                       if rule == "exact" else 'w:line="240" w:lineRule="auto"')
            ppr = f'<w:pPr><w:spacing w:before="0" w:after="0" {spacing}/>{rpr}</w:pPr>'
            body = f'<w:r>{rpr}<w:t xml:space="preserve">{tag} Hxyg</w:t></w:r>'
            if os.environ.get("EL_SUPER"):
                # A superscript and a subscript at the same base size. Word
                # draws both smaller and off the common baseline, so they are
                # the sharpest test of "align, THEN raise": getting the order
                # wrong misses by the raise, which is about a third of the size.
                # Anchor on `<w:sz ` with its space: `<w:sz` alone also matches
                # the start of `<w:szCs` and would insert the tag twice.
                srpr = rpr.replace('<w:sz ', '<w:vertAlign w:val="superscript"/><w:sz ')
                brpr = rpr.replace('<w:sz ', '<w:vertAlign w:val="subscript"/><w:sz ')
                body += f'<w:r>{srpr}<w:t>9</w:t></w:r><w:r>{brpr}<w:t>7</w:t></w:r>'
            if os.environ.get("EL_MIXED"):
                # A second VISIBLE run at a different size on the same line.
                # Word puts both on one baseline; an engine that places glyph
                # TOPS has to use each run's own ascent to get there.
                other = float(os.environ["EL_MIXED"])
                orpr = (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
                        f'<w:sz w:val="{round(other * 2)}"/>'
                        f'<w:szCs w:val="{round(other * 2)}"/></w:rPr>')
                body += f'<w:r>{orpr}<w:t xml:space="preserve"> Wm</w:t></w:r>'
            if os.environ.get("EL_BIG_SPACE"):
                # A trailing run of SPACES at a much larger size. S1045 says
                # such a fragment does not drive the line's height; whether it
                # drives the BASELINE of an exact line is a separate question,
                # and the conversion from 0.8 x line to a glyph top goes
                # through the largest fragment's ascent.
                big = float(os.environ["EL_BIG_SPACE"])
                brpr = (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
                        f'<w:sz w:val="{round(big * 2)}"/>'
                        f'<w:szCs w:val="{round(big * 2)}"/></w:rPr>')
                body += f'<w:r>{brpr}<w:t xml:space="preserve">   </w:t></w:r>'
            one = f'<w:p>{ppr}{body}</w:p>'
            if os.environ.get("EL_IN_CELL") == "1":
                # The same paragraph inside a one-cell table. A cell disables
                # grid snap and takes its own spacing reset, so it is a
                # different placement regime and has to be asked separately.
                one = ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>'
                       '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/>'
                       '<w:bottom w:val="single" w:sz="4" w:color="000000"/></w:tblBorders>'
                       '</w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>'
                       '<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>'
                       + one + '</w:tc></w:tr></w:tbl>')
            paras.append(one)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as o:
        for item in z.infolist():
            if item.filename == "word/document.xml":
                o.writestr(item, (head + "".join(paras) + sect
                                  + "</w:body></w:document>").encode("utf-8"))
            elif item.filename.startswith(("word/header", "word/footer")):
                continue
            else:
                data = z.read(item.filename)
                if item.filename == "word/_rels/document.xml.rels":
                    data = re.sub(
                        rb'<Relationship [^>]*Target="(header|footer)\d*\.xml"[^>]*/>', b"", data)
                if item.filename == "[Content_Types].xml":
                    data = re.sub(
                        rb'<Override [^>]*PartName="/word/(header|footer)\d*\.xml"[^>]*/>',
                        b"", data)
                o.writestr(item, data)
    top = int(re.search(r'w:top="(-?\d+)"',
                        re.search(r"<w:pgMar [^>]*/>", sect).group(0)).group(1)) / 20.0
    return plan, top


def main() -> int:
    stem = sys.argv[1] if len(sys.argv) > 1 else "ukframework"
    font = sys.argv[2] if len(sys.argv) > 2 else "Calibri"
    hosts = sorted(DOCS.glob(f"{stem}*.docx"))
    if not hosts:
        print(f"no host matching {stem}")
        return 1
    plan, top = build(hosts[0], font)

    import win32com.client
    pdf = str(OUT)[:-5] + ".pdf"
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    tops = {}
    try:
        d = app.Documents.Open(str(OUT), ReadOnly=True)
        # Ask Word for each line's TOP rather than accumulating the exact
        # heights. Inside a table the accumulation is wrong (borders and cell
        # margins are in the way) and the whole question is whether a cell
        # places the baseline the same way.
        for para in d.Paragraphs:
            text = para.Range.Text.replace("\r", "").replace("\x07", "").strip()
            if not text:
                continue
            rng = d.Range(para.Range.Start, para.Range.Start)
            tops.setdefault(text.split()[0], round(float(rng.Information(6)), 2))
        d.ExportAsFixedFormat(pdf, 17)
        d.Close(False)
    finally:
        app.Quit()

    import fitz
    page = fitz.open(pdf)[0]
    spans = [s for b in page.get_text("dict")["blocks"] for l in b.get("lines", [])
             for s in l["spans"] if s["text"].strip()]
    by_tag = {}
    for s in spans:
        tag = s["text"].split()[0]
        by_tag.setdefault(tag, s)

    where = "a table cell" if os.environ.get("EL_IN_CELL") == "1" else "the body"
    print(f"host {hosts[0].name} | font {font} | top margin {top:.2f}pt | in {where}")
    print(f"{'line':>5} {'size':>5} {'baseline':>9} {'line top':>9} "
          f"{'below top':>10} {'/ line':>7}")
    for line, size, tag in plan:
        s, box = by_tag.get(tag), tops.get(tag)
        if s is None or box is None:  # spilled to a later page
            continue
        below = s["origin"][1] - box
        print(f"{line:5g} {size:5g} {s['origin'][1]:9.2f} {box:9.2f} "
              f"{below:10.2f} {below / line:7.4f}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
