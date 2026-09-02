# -*- coding: utf-8 -*-
"""Read the BoundWidth sweep: which box property adds the constant?

Three numbers per arm, so the box and the advances can be told apart:

    design    the `hmtx` sum for the word out of the font file -- the pen
    BoundWidth  PowerPoint's line box
    engine    what the engine measured, from `--dump-layout`

An arm where `BoundWidth - design` is ~0 and one where it is ~3pt, with only
one property between them, names the property.

★Must not overlap the renderer or another COM session.

    python tools/metrics/read_pptx_boundwidth_com.py
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import win32com.client
from fontTools.ttLib import TTFont

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
from pptx_gdi_face_audit import registry_file  # noqa: E402

DECK = REPO / "pipeline_data" / "pptx_probes" / "boundwidth" / "boundwidth.pptx"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"
WORD = "Yellow"
FACE = "Arial"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def design_width(size: float, bold: bool = False, italic: bool = False) -> float | None:
    name = FACE + (" Bold" if bold and not italic else "")
    path = registry_file(name) or registry_file(FACE)
    if path is None:
        return None
    font = TTFont(str(path))
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    total = 0
    for ch in WORD:
        g = cmap.get(ord(ch))
        if g is None:
            return None
        total += hmtx[g][0]
    return total / upm * size


def engine_widths() -> dict[tuple[float, float], float]:
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "l.json"
        subprocess.run([str(EXE), str(DECK), str(Path(td) / "s"), "150",
                        f"--dump-layout={out}"],
                       capture_output=True, timeout=600, check=False)
        if not out.exists():
            return {}
        dump = json.loads(out.read_text(encoding="utf-8"))
    got: dict[tuple[float, float], float] = {}
    for slide in dump.get("slides", []):
        for sh in slide.get("shapes", []):
            for p in (sh.get("content") or {}).get("paragraphs") or []:
                text = "".join(r.get("text", "") for r in p.get("runs", []))
                if text.strip() == WORD:
                    w = (p.get("line_widths") or [None])[0]
                    if w is not None:
                        got[(round(sh["x"], 1), round(sh["y"], 1))] = w
    return got


def main() -> None:
    if not DECK.exists():
        sys.exit(f"{DECK} is not there -- run gen_pptx_boundwidth.py first")
    mine = engine_widths()

    app = win32com.client.Dispatch("PowerPoint.Application")
    rows = []
    try:
        pres = app.Presentations.Open(str(DECK.resolve()), WithWindow=False)
        try:
            slide = pres.Slides(1)
            labels, arms = [], []
            for i in range(1, slide.Shapes.Count + 1):
                sh = slide.Shapes(i)
                if not sh.HasTextFrame or not sh.TextFrame.HasText:
                    continue
                text = sh.TextFrame.TextRange.Text.strip()
                if text == WORD:
                    arms.append(sh)
                else:
                    labels.append((sh.Left, sh.Top, text))
            for sh in arms:
                tr = sh.TextFrame.TextRange
                above = [(t, n) for (x, t, n) in labels
                         if abs(x - sh.Left) < 1.0 and t < sh.Top]
                label = max(above)[1] if above else f"x={sh.Left:.0f}"
                rows.append((label, sh.Left, sh.Top, tr.Lines(1).BoundWidth,
                             float(tr.Font.Size), bool(tr.Font.Bold == -1),
                             bool(tr.Font.Italic == -1)))
        finally:
            pres.Saved = True
            pres.Close()
    finally:
        app.Quit()

    print(f"{WORD!r} in {FACE}\n")
    print(f"{'arm':<10}{'size':>6}{'design':>9}{'BoundWidth':>12}{'engine':>9}"
          f"{'box-design':>12}{'engine-design':>15}")
    for label, left, top, bw, size, bold, italic in sorted(rows, key=lambda r: (r[2], r[1])):
        d = design_width(size, bold, italic)
        e = mine.get((round(left, 1), round(top, 1)))
        if d is None:
            print(f"{label:<10}{size:>6.0f}   (no font file for {FACE})")
            continue
        eng = f"{e:9.2f}" if e is not None else "        -"
        gap = f"{e - d:+15.2f}" if e is not None else "              -"
        print(f"{label:<10}{size:>6.0f}{d:>9.2f}{bw:>12.2f}{eng}{bw - d:>+12.2f}{gap}")


if __name__ == "__main__":
    main()
