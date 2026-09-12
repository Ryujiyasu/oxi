# -*- coding: utf-8 -*-
"""How tall is one line — Word's pitch against this engine's.

A JA document that matches Word paragraph for paragraph still puts 41 lines on
a page where Word puts 57, because its line pitch is 13.5pt against Word's
12.3 — about a tenth too tall, every line, all the way down. Capacity is right
(`_pb_line_capacity.py`, `_pb_page_capacity.py` both agree everywhere); the
line itself is not.

The document's shape is a `w:docGrid` that states a `linePitch` and NO
`w:type`, a CJK font for east-asian text with a Latin font for ascii, and a
1.15 line multiple. This builds that shape and takes it apart one setting at a
time: the pitch is read as the distance between consecutive baselines, from
Word's own PDF and from the engine's dump.

    python tools/metrics/_pb_line_pitch.py
    LP_GRID=typed LP_LINE=240 python tools/metrics/_pb_line_pitch.py
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
OUT = REPO / "tests" / "fixtures" / "line_pitch"

# A document with no styles part falls back to the ENGINE's own default face
# and size, and the first sweep did exactly that: the debug line reported
# 游明朝 at 10.5pt while the probe asked for MS Mincho at 10pt, so Word and the
# engine were being compared on different fonts. The styles part below pins
# docDefaults to the probe's own choice.
CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                 '</Types>')


def styles(ascii_font: str, ea_font: str, size: float) -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{ascii_font}" w:eastAsia="{ea_font}" '
            f'w:hAnsi="{ascii_font}" w:cs="{ascii_font}"/>'
            f'<w:sz w:val="{round(size * 2)}"/><w:szCs w:val="{round(size * 2)}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/><w:qFormat/></w:style>'
            '</w:styles>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
             '</Relationships>')
DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>')


def document(n: int, grid: str, pitch: int, ascii_font: str, ea_font: str,
             size: float, line: int, rule: str, text: str) -> str:
    # `LP_HINT=1` adds w:hint="eastAsia". Word picks the east-asian face by
    # CHARACTER whatever the hint says; the engine was seen resolving a line of
    # kanji to the ascii face without it, so the hint has to be a lever rather
    # than an assumption.
    hint = ' w:hint="eastAsia"' if os.environ.get("LP_HINT", "1") == "1" else ""
    rpr = (f'<w:rPr><w:rFonts w:ascii="{ascii_font}" w:eastAsia="{ea_font}" '
           f'w:hAnsi="{ascii_font}"{hint}/>'
           f'<w:sz w:val="{round(size * 2)}"/><w:szCs w:val="{round(size * 2)}"/></w:rPr>')
    ppr = ('<w:pPr>'
           f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="{rule}"/>'
           f'{rpr}</w:pPr>')
    # The counter used to be appended as ASCII digits, which put a LATIN
    # fragment on every line and made "pure CJK" impossible to ask for. Kanji
    # numerals keep the line east-asian; `LP_COUNT=0` drops it entirely.
    digits = "〇一二三四五六七八九"
    def tag(i: int) -> str:
        if os.environ.get("LP_COUNT", "1") != "1":
            return ""
        return "".join(digits[int(c)] for c in f"{i + 1:03d}")
    # `LP_MIX_SIZE` adds a LATIN run at its own size. With a Latin box TALLER
    # than the CJK one it separates the two candidate rules: "max of each
    # fragment's whole box" against "max ascent plus max descent across
    # fragments", which differ only when the taller box is not the one with the
    # deeper descent.
    mix_size = float(os.environ.get("LP_MIX_SIZE", "0"))
    extra = ""
    if mix_size > 0:
        mrpr = (f'<w:rPr><w:rFonts w:ascii="{ascii_font}" w:hAnsi="{ascii_font}"/>'
                f'<w:sz w:val="{round(mix_size * 2)}"/>'
                f'<w:szCs w:val="{round(mix_size * 2)}"/></w:rPr>')
        extra = f'<w:r>{mrpr}<w:t xml:space="preserve"> Ag</w:t></w:r>'
    paras = "".join(f'<w:p>{ppr}<w:r>{rpr}<w:t>{text}{tag(i)}</w:t></w:r>{extra}</w:p>'
                    for i in range(n))
    if grid == "none":
        dg = ""
    elif grid == "notype":
        # The shape the failing document carries: a pitch with no type at all.
        dg = f'<w:docGrid w:linePitch="{pitch}"/>'
    else:
        dg = f'<w:docGrid w:type="{grid}" w:linePitch="{pitch}"/>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
            f'w:header="0" w:footer="0" w:gutter="0"/>{dg}</w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{paras}{sect}</w:body></w:document>')


def write(path: Path, xml: str, style_xml: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", style_xml)
        z.writestr("word/document.xml", xml)


def oxi_pitch(path: Path):
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "l.json"
        subprocess.run([str(GDI), str(path), str(Path(tmp) / "p"), "150",
                        f"--dump-layout={dump}"], capture_output=True)
        if not dump.is_file():
            return None, 0
        data = json.loads(dump.read_text(encoding="utf-8"))
    pages = data.get("pages", [])
    if not pages:
        return None, 0
    ys = sorted({round(float(e["y"]), 2) for e in pages[0].get("elements", [])
                 if e.get("type") == "text" and e.get("text")})
    gaps = [round(b - a, 2) for a, b in zip(ys, ys[1:])]
    return (max(set(gaps), key=gaps.count) if gaps else None), len(ys)


def word_pitch(app, path: Path):
    import fitz
    pdf = str(path)[:-5] + ".pdf"
    doc = app.Documents.Open(str(path), False, True)
    try:
        doc.ExportAsFixedFormat(pdf, 17)
    finally:
        doc.Close(False)
    page = fitz.open(pdf)[0]
    ys = sorted({round(line["bbox"][1], 2) for b in page.get_text("dict")["blocks"]
                 for line in b.get("lines", [])})
    gaps = [round(b - a, 2) for a, b in zip(ys, ys[1:])]
    return (max(set(gaps), key=gaps.count) if gaps else None), len(ys)


def main() -> int:
    grid = os.environ.get("LP_GRID", "notype")
    pitch = int(os.environ.get("LP_PITCH", "360"))
    ascii_font = os.environ.get("LP_ASCII", "Arial")
    ea_font = os.environ.get("LP_EA", "MS Mincho")
    size = float(os.environ.get("LP_SIZE", "10"))
    line = int(os.environ.get("LP_LINE", "276"))
    rule = os.environ.get("LP_RULE", "auto")
    text = os.environ.get("LP_TEXT", "本文の行")
    n = int(os.environ.get("LP_N", "40"))

    at = OUT / f"pitch_{grid}{pitch}_{ascii_font}_{ea_font}_s{size:g}_{line}{rule}.docx".replace(" ", "")
    write(at, document(n, grid, pitch, ascii_font, ea_font, size, line, rule, text),
          styles(ascii_font, ea_font, size))

    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    for attr, value in (("Visible", False), ("DisplayAlerts", False)):
        try:
            setattr(app, attr, value)
        except Exception:  # noqa: BLE001
            pass
    try:
        wp, wn = word_pitch(app, at)
    finally:
        app.Quit()
    op, on = oxi_pitch(at)
    same = wp is not None and op is not None and abs(wp - op) < 0.05
    print(f"grid={grid}:{pitch} ascii={ascii_font} ea={ea_font} size={size:g} "
          f"line={line}{rule} text={text!r}")
    print(f"    word pitch {wp} ({wn} lines)   oxi pitch {op} ({on} lines)"
          f"{'' if same else '   <<<'}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
