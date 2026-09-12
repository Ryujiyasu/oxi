# -*- coding: utf-8 -*-
"""The first paragraph whose LINE COUNT differs — the cause, not the symptom.

Half the Phase-1 failures have their first page slip on an ordinary paragraph
with nothing special about it. That paragraph is innocent: a page slip means
some earlier paragraph took one line too many or too few, and everything after
it carries the error. Asking which paragraph SLIPPED therefore names a victim.

So ask a different question. Word can report how many lines it gives each
paragraph, and this engine's layout dump says the same. The first paragraph
where the two disagree is where the document actually goes wrong.

    python tools/metrics/_p1_first_line_diff.py <docx path> [...]

Prints, per document, the first disagreement and the few around it, with the
paragraph's own properties so the construct is visible without reading text.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
GDI = REPO / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"


def word_lines(app, path: str) -> list:
    """(text prefix, line count) per paragraph, as Word lays it out."""
    doc = app.Documents.Open(path, False, True)
    try:
        out = []
        for para in doc.Paragraphs:
            rng = para.Range
            text = rng.Text.replace("\r", "").replace("\x07", "")
            # 1 = wdStatisticLines
            try:
                n = int(rng.ComputeStatistics(1))
            except Exception:  # noqa: BLE001
                n = -1
            out.append((text, n))
        return out
    finally:
        doc.Close(False)


def oxi_lines(path: str) -> list:
    """(paragraph index, line count) as this engine lays it out."""
    with tempfile.TemporaryDirectory() as tmp:
        dump = Path(tmp) / "l.json"
        subprocess.run([str(GDI), path, str(Path(tmp) / "p"), "150",
                        f"--dump-layout={dump}"], capture_output=True)
        if not dump.is_file():
            return []
        data = json.loads(dump.read_text(encoding="utf-8"))
    seen: dict[int, set] = {}
    for page in data.get("pages", []):
        for e in page.get("elements", []):
            if e.get("type") != "text" or not e.get("text"):
                continue
            pi = e.get("para_idx")
            if pi is None:
                continue
            seen.setdefault(pi, set()).add(round(float(e["y"]), 1))
    return [(pi, len(ys)) for pi, ys in sorted(seen.items())]


def properties(path: str) -> list:
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8", "replace")
    body = xml[xml.index("<w:body>"):]
    out, depth_tbl, depth_txbx = [], 0, 0
    for m in re.finditer(r"<(/?)w:(tbl|txbxContent|p)([ />])", body):
        closing, name, _ = m.groups()
        if name == "tbl":
            depth_tbl += -1 if closing else 1
        elif name == "txbxContent":
            depth_txbx += -1 if closing else 1
        elif name == "p" and not closing:
            end = body.find("</w:p>", m.end())
            chunk = body[m.start():end if end > 0 else m.end()]
            kinds = []
            if depth_txbx > 0:
                kinds.append("textbox")
            if depth_tbl > 0:
                kinds.append("table")
            for tag, label in (("w:numPr", "numbering"), ("w:tabs", "tabs"),
                               ("w:ind ", "indent"), ("w:jc ", "justify"),
                               ("w:spacing", "spacing"), ("w:drawing", "drawing"),
                               ("instrText", "field"), ("w:br", "break"),
                               ("w:noBreakHyphen", "noBreakHyphen"),
                               ("w:sz ", "size")):
                if tag in chunk:
                    kinds.append(label)
            out.append(kinds or ["plain"])
    return out


def main() -> int:
    paths = sys.argv[1:]
    if not paths:
        print(f"usage: {Path(sys.argv[0]).name} <docx path> [...]")
        return 2
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    for attr, value in (("Visible", False), ("DisplayAlerts", False)):
        try:
            setattr(app, attr, value)
        except Exception:  # noqa: BLE001
            pass
    try:
        for path in paths:
            path = str(Path(path).resolve())
            print(f"\n=== {Path(path).name}")
            wl = word_lines(app, path)
            ol = dict(oxi_lines(path))
            props = properties(path)
            # The engine numbers only paragraphs it drew; a paragraph with no
            # visible glyph is absent, and that is not a disagreement.
            first = None
            for i, (text, n) in enumerate(wl):
                if not text.strip():
                    continue
                got = ol.get(i)
                if got is None or n < 0:
                    continue
                if got != n:
                    first = i
                    break
            if first is None:
                print("  every drawn paragraph takes the same number of lines")
                continue
            print(f"{'para':>6} {'word':>5} {'oxi':>5}  properties")
            for i in range(max(0, first - 3), min(len(wl), first + 4)):
                text, n = wl[i]
                got = ol.get(i)
                mark = " <<<" if i == first else ""
                kinds = "+".join(props[i]) if i < len(props) else "?"
                print(f"{i:6} {n:5} {str(got):>5}  {kinds}{mark}")
    finally:
        app.Quit()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
