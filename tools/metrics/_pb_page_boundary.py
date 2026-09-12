# -*- coding: utf-8 -*-
"""At what spacer height does a page give up its last line?

Removing the half-point snap from the no-grid line height moved
`footer_spacing/empty.docx` from 11 pages to 10, and that fixture is a
recorded Word measurement. But it is also a knife edge by construction: an
`exact` spacer sized to leave room for exactly one more body line, so a change
of 0.03pt in that body line decides the page.

A single knife edge cannot say which line height is right — it only says the
two disagree. So this sweeps the spacer instead: the value at which the mark
crosses to the next page is a direct reading of the space the body line takes,
and Word's crossing and the engine's can be compared as numbers.

    python tools/metrics/_pb_page_boundary.py
    PB_LO=5700 PB_HI=6000 PB_STEP=20 python tools/metrics/_pb_page_boundary.py
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
SRC = REPO / "tests" / "fixtures" / "footer_spacing" / "empty.docx"
OUT = REPO / "tests" / "fixtures" / "page_boundary"


def build(spacer_tw: int) -> Path:
    """The fixture's own shape with one case and a swept spacer."""
    z = zipfile.ZipFile(SRC)
    xml = z.read("word/document.xml").decode("utf-8")
    head = xml[: xml.index("<w:body>") + len("<w:body>")]
    tail = xml[xml.rindex("</w:body>"):]
    import re
    sect = re.findall(r"<w:sectPr[ >].*?</w:sectPr>", xml, re.S)[-1]
    rpr = ('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
           '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>')

    def para(text: str, line: int, rule: str) -> str:
        return ('<w:p><w:pPr>'
                f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="{rule}"/>'
                f'{rpr}</w:pPr><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>')

    body = (para("HEAD", 240, "auto")
            + para("SPACER", spacer_tw, "exact")
            + para("MARK", 240, "auto"))
    OUT.mkdir(parents=True, exist_ok=True)
    at = OUT / f"boundary_{spacer_tw}.docx"
    with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as o:
        for item in z.infolist():
            if item.filename == "word/document.xml":
                o.writestr(item, (head + body + sect + tail).encode("utf-8"))
            else:
                o.writestr(item, z.read(item.filename))
    return at


def oxi_mark_page(path: Path, env_extra=None) -> int:
    env = dict(os.environ)
    env.update(env_extra or {})
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "l.json"
        subprocess.run([str(GDI), str(path), str(Path(tmp) / "p"), "150",
                        f"--dump-layout={dump}"], capture_output=True, env=env)
        if not dump.is_file():
            return -1
        data = json.loads(dump.read_text(encoding="utf-8"))
    for i, page in enumerate(data.get("pages", []), 1):
        if any(e.get("text") == "MARK" for e in page.get("elements", [])):
            return i
    return -1


def word_mark_page(app, path: Path) -> int:
    doc = app.Documents.Open(str(path), False, True)
    try:
        for para in doc.Paragraphs:
            if para.Range.Text.replace("\r", "").strip() == "MARK":
                rng = doc.Range(para.Range.Start, para.Range.Start)
                return int(rng.Information(3))
        return -1
    finally:
        doc.Close(False)


def main() -> int:
    lo = int(os.environ.get("PB_LO", "5700"))
    hi = int(os.environ.get("PB_HI", "6020"))
    step = int(os.environ.get("PB_STEP", "20"))

    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    for attr, value in (("Visible", False), ("DisplayAlerts", False)):
        try:
            setattr(app, attr, value)
        except Exception:  # noqa: BLE001
            pass
    print(f"{'spacer tw':>10} {'pt':>8} {'word':>5} {'oxi':>5} {'snapped':>8}")
    flips = {}
    try:
        for tw in range(lo, hi + 1, step):
            at = build(tw)
            w = word_mark_page(app, at)
            o = oxi_mark_page(at)
            s = oxi_mark_page(at, {"OXI_S1362_DISABLE": "1"})
            for name, v in (("word", w), ("oxi", o), ("snapped", s)):
                if v > 1 and name not in flips:
                    flips[name] = tw
            print(f"{tw:10} {tw / 20:8.2f} {w:5} {o:5} {s:8}")
    finally:
        app.Quit()
    print("\nfirst spacer that pushes MARK off page 1:")
    for name in ("word", "oxi", "snapped"):
        print(f"  {name:8} {flips.get(name, '(none in range)')}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
