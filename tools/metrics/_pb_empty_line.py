# -*- coding: utf-8 -*-
"""How tall is an EMPTY paragraph — Word's answer against this engine's.

Removing the half-point snap from no-grid lines put every measured BODY pitch
on Word (Arial, Times New Roman and MS Mincho all within 0.005pt), and moved
one recorded fixture's page boundary by 0.25pt. The line that moved it is not
body text at all: it is the empty first paragraph of a footer, whose height
comes down a different path and which the snap was quietly rounding up.

An empty paragraph draws nothing, so its height cannot be read from its own
position. This puts N of them between two visible lines and sweeps N: the gap
grows by exactly one empty line per step, so the SLOPE is the height, with the
surrounding paragraphs' own heights cancelling out.

    python tools/metrics/_pb_empty_line.py
    EL_FONT=Arial EL_SIZE=10 python tools/metrics/_pb_empty_line.py
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
OUT = REPO / "tests" / "fixtures" / "empty_line"

CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                 '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
             '</Relationships>')
DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>')


def styles(font: str, size: float) -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{round(size * 2)}"/><w:szCs w:val="{round(size * 2)}"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>')


def document(n_empty: int, font: str, size: float, empty_size: float) -> str:
    def rpr(fs: float) -> str:
        return (f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
                f'<w:sz w:val="{round(fs * 2)}"/><w:szCs w:val="{round(fs * 2)}"/></w:rPr>')
    ppr = ('<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
           f'{rpr(size)}</w:pPr>')
    eppr = ('<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'{rpr(empty_size)}</w:pPr>')
    body = f'<w:p>{ppr}<w:r>{rpr(size)}<w:t>TOP</w:t></w:r></w:p>'
    body += f'<w:p>{eppr}</w:p>' * n_empty
    body += f'<w:p>{ppr}<w:r>{rpr(size)}<w:t>BOTTOM</w:t></w:r></w:p>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
            'w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}{sect}</w:body></w:document>')


def write(path: Path, xml: str, style_xml: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", style_xml)
        z.writestr("word/document.xml", xml)


def gap_word(app, path: Path) -> float:
    import fitz
    pdf = str(path)[:-5] + ".pdf"
    doc = app.Documents.Open(str(path), False, True)
    try:
        doc.ExportAsFixedFormat(pdf, 17)
    finally:
        doc.Close(False)
    tops = {}
    for b in fitz.open(pdf)[0].get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            text = "".join(s["text"] for s in line["spans"]).strip()
            if text in ("TOP", "BOTTOM"):
                tops[text] = line["bbox"][1]
    return round(tops["BOTTOM"] - tops["TOP"], 3) if len(tops) == 2 else float("nan")


def gap_oxi(path: Path, env_extra=None) -> float:
    env = dict(os.environ)
    env.update(env_extra or {})
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "l.json"
        subprocess.run([str(GDI), str(path), str(Path(tmp) / "p"), "150",
                        f"--dump-layout={dump}"], capture_output=True, env=env)
        if not dump.is_file():
            return float("nan")
        data = json.loads(dump.read_text(encoding="utf-8"))
    tops = {}
    for page in data.get("pages", []):
        for e in page.get("elements", []):
            if e.get("text") in ("TOP", "BOTTOM"):
                tops[e["text"]] = float(e["y"])
    return round(tops["BOTTOM"] - tops["TOP"], 3) if len(tops) == 2 else float("nan")


def main() -> int:
    font = os.environ.get("EL_FONT", "Arial")
    size = float(os.environ.get("EL_SIZE", "10"))
    empty_size = float(os.environ.get("EL_EMPTY_SIZE", str(size)))
    counts = [int(x) for x in os.environ.get("EL_N", "0,1,2,4,8").split(",")]

    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    for attr, value in (("Visible", False), ("DisplayAlerts", False)):
        try:
            setattr(app, attr, value)
        except Exception:  # noqa: BLE001
            pass
    rows = []
    try:
        for n in counts:
            at = OUT / f"empty_{font}_{size:g}_{empty_size:g}_{n}.docx".replace(" ", "")
            write(at, document(n, font, size, empty_size), styles(font, size))
            rows.append((n, gap_word(app, at), gap_oxi(at),
                         gap_oxi(at, {"OXI_S1362_DISABLE": "1"})))
    finally:
        app.Quit()

    print(f"font={font} size={size:g} empty run size={empty_size:g}")
    print(f"{'empties':>8} {'word gap':>9} {'oxi':>9} {'snapped':>9}")
    for n, w, o, s in rows:
        print(f"{n:8} {w:9.3f} {o:9.3f} {s:9.3f}")
    if len(rows) >= 2:
        def slope(i):
            (n0, *v0), (n1, *v1) = rows[0], rows[-1]
            return (v1[i] - v0[i]) / (n1 - n0)
        print(f"\nheight of one empty paragraph: word {slope(0):.4f}  "
              f"oxi {slope(1):.4f}  snapped {slope(2):.4f}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
