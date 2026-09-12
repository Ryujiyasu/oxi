# -*- coding: utf-8 -*-
"""How many lines does one page hold — Word's answer against this engine's.

`_pb_line_capacity.py` settled the horizontal question: plain Japanese wraps at
exactly the same character on both sides, for every line of a six-line
paragraph and under every setting swept. So the line the engine loses on real
documents is not lost across the line, it is lost down the page.

This asks the vertical question the same way: a document of M one-line
paragraphs, and the paragraph at which each side starts page 2 IS its page
capacity. One line of difference is the whole bug.

    python tools/metrics/_pb_page_capacity.py
    PC_GRID=0 PC_SIZE=9 PC_SPACE=300 python tools/metrics/_pb_page_capacity.py

Sweeps what could move it — the grid pitch, the line rule, the font size,
space before and after — because a capacity that is right bare and wrong with
spacing names the term.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
GDI = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"
OUT = REPO / "tests" / "fixtures" / "page_capacity"
FONT = "ＭＳ 明朝"

CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
             '</Relationships>')


def document(n: int, grid: int, size: float, before: int, after: int,
             line: int, rule: str) -> str:
    rpr = (f'<w:rPr><w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/>'
           f'<w:sz w:val="{round(size * 2)}"/><w:szCs w:val="{round(size * 2)}"/></w:rPr>')
    # schema order inside pPr: spacing, then jc.
    ppr = ('<w:pPr>'
           f'<w:spacing w:before="{before}" w:after="{after}" '
           f'w:line="{line}" w:lineRule="{rule}"/>'
           f'<w:jc w:val="both"/>{rpr}</w:pPr>')
    one = [f'<w:p>{ppr}<w:r>{rpr}<w:t>行{i + 1:03d}あいうえお</w:t></w:r></w:p>'
           for i in range(n)]
    if os.environ.get("PC_IN_CELL") == "1":
        # The same paragraphs inside a single-cell table. A cell carries its own
        # margins and its own line-height rule, so a page capacity that is right
        # in the body can still be wrong here — and half the failures whose
        # first slip is "plain" sit in documents full of tables.
        border = "".join(f'<w:{side} w:val="single" w:sz="4" w:color="000000"/>'
                         for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
        one = ['<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>'
               f'<w:tblBorders>{border}</w:tblBorders></w:tblPr>'
               '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
               '<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>'
               + "".join(one) + '</w:tc></w:tr></w:tbl>']
    paras = "".join(one)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
            'w:header="0" w:footer="0" w:gutter="0"/>'
            + (f'<w:docGrid w:type="lines" w:linePitch="{grid}"/>' if grid else '')
            + '</w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{paras}{sect}</w:body></w:document>')


def write(path: Path, xml: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", xml)


def oxi_page_one(path: Path) -> int:
    """Paragraphs the engine put on page 1."""
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "l.json"
        subprocess.run([str(GDI), str(path), str(Path(tmp) / "p"), "150",
                        f"--dump-layout={dump}"], capture_output=True)
        if not dump.is_file():
            return -1
        data = json.loads(dump.read_text(encoding="utf-8"))
    pages = data.get("pages", [])
    if not pages:
        return -1
    # Count LINES, not paragraph indices: inside a table every paragraph of a
    # cell reports the table's own block index, so counting indices there
    # returns 1 however many lines are on the page.
    return len({round(float(e["y"]), 1) for e in pages[0].get("elements", [])
                if e.get("type") == "text" and e.get("text")})


def word_page_one(app, path: Path) -> int:
    doc = app.Documents.Open(str(path), False, True)
    try:
        n = 0
        for para in doc.Paragraphs:
            rng = doc.Range(para.Range.Start, para.Range.Start)
            # 3 = wdActiveEndPageNumber, on a collapsed start (the R30 rule).
            if int(rng.Information(3)) == 1:
                n += 1
            else:
                break
        return n
    finally:
        doc.Close(False)


def main() -> int:
    grid = int(os.environ.get("PC_GRID", "360"))
    size = float(os.environ.get("PC_SIZE", "10.5"))
    before = int(os.environ.get("PC_BEFORE", "0"))
    after = int(os.environ.get("PC_AFTER", "0"))
    line = int(os.environ.get("PC_LINE", "240"))
    rule = os.environ.get("PC_RULE", "auto")
    n = int(os.environ.get("PC_N", "90"))

    where = "cell" if os.environ.get("PC_IN_CELL") == "1" else "body"
    at = OUT / f"page_{where}_g{grid}_s{size:g}_b{before}_a{after}_l{line}{rule}.docx"
    write(at, document(n, grid, size, before, after, line, rule))

    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    for attr, value in (("Visible", False), ("DisplayAlerts", False)):
        try:
            setattr(app, attr, value)
        except Exception:  # noqa: BLE001
            pass
    try:
        w = word_page_one(app, at)
    finally:
        app.Quit()
    o = oxi_page_one(at)
    mark = "" if w == o else "   <<<"
    print(f"{where} grid={grid} size={size:g} before={before} after={after} "
          f"line={line}{rule}  ->  page 1 holds  word {w}  oxi {o}{mark}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
