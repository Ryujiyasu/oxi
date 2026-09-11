# -*- coding: utf-8 -*-
"""How much smaller is a superscript, exactly?

This engine draws one at 0.583 of the base size. Word's own PDF reports 8.04
for a 12pt base, which is 0.67, but the reported size is rounded and a ratio
read off it wanders between 0.63 and 0.67 across the sizes. So this does not
read the size at all: it puts the SAME character at the base size and again as
a superscript, and compares the two glyph boxes. A box ratio is a size ratio,
with no rounding in the way.

    python tools/metrics/_pb_superscript_size.py [host-stem] [font]

Also reports the rise and the drop as fractions of the base size, measured the
same way, because a wrong size and a wrong rise look alike on the page.
"""
import os
import re
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
DOCS = REPO / "tools" / "golden-test" / "documents" / "docx"
OUT = REPO / "pipeline_data" / "_pb_superscript_size.docx"

SIZES = [float(x) for x in os.environ.get("SS_SIZES", "8,10,12,16,20,24").split(",")]
# A tall flat-topped digit with no descender, so the box is the cap height and
# both copies clip the same way.
MARK = "8"


def build(host: Path, font: str) -> list:
    z = zipfile.ZipFile(host)
    doc = z.read("word/document.xml").decode("utf-8")
    head = doc[: doc.index("<w:body>") + len("<w:body>")]
    sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", doc, re.S)[-1]
    sect = re.sub(r"<w:(header|footer)Reference[^>]*/>", "", sect)

    def rpr(size: float, align: str = "") -> str:
        tag = f'<w:vertAlign w:val="{align}"/>' if align else ""
        return (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>{tag}'
                f'<w:sz w:val="{round(size * 2)}"/>'
                f'<w:szCs w:val="{round(size * 2)}"/></w:rPr>')

    paras, plan = [], []
    for size in SIZES:
        tag = f"Z{int(size):02d}"
        plan.append((size, tag))
        # tag, then the mark three times: normal, superscript, subscript.
        paras.append(
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
            'w:lineRule="auto"/></w:pPr>'
            f'<w:r>{rpr(size)}<w:t xml:space="preserve">{tag} </w:t></w:r>'
            f'<w:r>{rpr(size)}<w:t>{MARK}</w:t></w:r>'
            f'<w:r>{rpr(size)}<w:t xml:space="preserve"> </w:t></w:r>'
            f'<w:r>{rpr(size, "superscript")}<w:t>{MARK}</w:t></w:r>'
            f'<w:r>{rpr(size)}<w:t xml:space="preserve"> </w:t></w:r>'
            f'<w:r>{rpr(size, "subscript")}<w:t>{MARK}</w:t></w:r></w:p>')

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
    return plan


def main() -> int:
    stem = sys.argv[1] if len(sys.argv) > 1 else "ukframework"
    font = sys.argv[2] if len(sys.argv) > 2 else "Calibri"
    hosts = sorted(DOCS.glob(f"{stem}*.docx"))
    if not hosts:
        print(f"no host matching {stem}")
        return 1
    plan = build(hosts[0], font)

    import win32com.client
    pdf = str(OUT)[:-5] + ".pdf"
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(str(OUT), ReadOnly=True)
        d.ExportAsFixedFormat(pdf, 17)
        d.Close(False)
    finally:
        app.Quit()

    import fitz
    page = fitz.open(pdf)[0]
    # One character per span is not guaranteed, so walk the raw characters.
    rows = {}
    for b in page.get_text("rawdict")["blocks"]:
        for line in b.get("lines", []):
            for span in line["spans"]:
                for ch in span["chars"]:
                    if ch["c"].strip():
                        rows.setdefault(round(line["bbox"][1], 1), []).append(
                            (ch["bbox"][0], ch["c"], ch["bbox"], ch["origin"][1], span["size"]))
    print(f"host {hosts[0].name} | font {font} | mark {MARK!r}")
    print(f"{'base':>5} {'box':>7} {'sup box':>8} {'ratio':>7} "
          f"{'rise':>7} {'/base':>7} {'drop':>7} {'/base':>7}")
    for key in sorted(rows):
        chars = sorted(rows[key])
        marks = [c for c in chars if c[1] == MARK]
        if len(marks) != 3:
            continue
        base, sup, sub = marks
        height = lambda c: c[2][3] - c[2][1]  # noqa: E731
        size = height(base)
        if size <= 0:
            continue
        rise = base[3] - sup[3]
        drop = sub[3] - base[3]
        # The base size is recoverable from the run: the normal mark is set at it.
        pt = base[4]
        print(f"{pt:5.1f} {size:7.3f} {height(sup):8.3f} {height(sup)/size:7.4f} "
              f"{rise:7.2f} {rise/pt:7.4f} {drop:7.2f} {drop/pt:7.4f}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
