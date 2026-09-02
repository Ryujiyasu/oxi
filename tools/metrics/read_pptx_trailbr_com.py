# -*- coding: utf-8 -*-
"""Ask PowerPoint how many lines each trailing-break arm has, and how tall.

`Lines.Count` answers the break; `BoundHeight` answers whether the line is
reserved anyway. Both are needed: a trailing break that is not counted but is
still given its height would move every centred shape that carries one, and an
implementation that only fixes the count would move them the other way.

The engine's own answer for the same deck is printed beside it, from
`--dump-layout`, so the two are read together rather than a page apart.

★Must not run while the renderer is producing PNGs
(`pptx_com_render_must_not_overlap`).

    python tools/metrics/read_pptx_trailbr_com.py
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
DECK = REPO / "pipeline_data" / "pptx_probes" / "trailbr" / "trailbr.pptx"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"


def engine_lines() -> dict[float, int]:
    """The engine's line count per text shape, keyed by the shape's left edge."""
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "layout.json"
        subprocess.run(
            [str(EXE), str(DECK), str(Path(td) / "slide"), "150", f"--dump-layout={out}"],
            capture_output=True, timeout=600, check=False)
        if not out.exists():
            return {}
        dump = json.loads(out.read_text(encoding="utf-8"))
    got: dict[tuple[float, float], int] = {}
    for slide in dump.get("slides", []):
        for sh in slide.get("shapes", []):
            paras = (sh.get("content") or {}).get("paragraphs") or []
            text = "".join(r.get("text", "")
                           for p in paras for r in (p.get("runs") or []))
            # ★The label shape shares its arm's left edge, so keying on x alone
            # let the label overwrite the arm and every arm read "1 line" --
            # a probe that answered the same number for seven different inputs,
            # which is the shape a broken reader takes.
            if text.strip().startswith(("A_", "B_", "C_", "D_", "E_", "F_", "G_")):
                continue
            n = sum(len(p.get("line_x_offsets") or []) for p in paras)
            if n:
                got[(round(sh["x"], 1), round(sh["y"], 1))] = n
    return got


def main() -> None:
    if not DECK.exists():
        sys.exit(f"{DECK} is not there -- run gen_pptx_trailbr.py first")
    mine = engine_lines()

    app = win32com.client.Dispatch("PowerPoint.Application")
    rows: list[tuple[str, int, float, tuple[float, float]]] = []
    try:
        pres = app.Presentations.Open(str(DECK.resolve()), WithWindow=False)
        try:
            slide = pres.Slides(1)
            # The label shapes sit above the arms and share their left edge, so
            # each arm is named by the shape 228600 EMU above it.
            # ★Two rows now share each column, so a label is found by the one
            # sitting just ABOVE the arm, not by the left edge alone. Keying on
            # x only is what made every arm read the same number once the
            # wrapped row was added.
            labels: list[tuple[float, float, str]] = []
            arms = []
            for i in range(1, slide.Shapes.Count + 1):
                sh = slide.Shapes(i)
                if not sh.HasTextFrame or not sh.TextFrame.HasText:
                    continue
                text = sh.TextFrame.TextRange.Text
                if re.match(r"[A-G](W)?_", text.strip()):
                    labels.append((sh.Left, sh.Top, text.strip()))
                else:
                    arms.append(sh)
            for sh in arms:
                tr = sh.TextFrame.TextRange
                above = [(t, n) for (x, t, n) in labels
                         if abs(x - sh.Left) < 1.0 and t < sh.Top]
                label = max(above)[1] if above else f"x={sh.Left:.0f} y={sh.Top:.0f}"
                rows.append((label, tr.Lines().Count, tr.BoundHeight,
                             (sh.Left, sh.Top)))
        finally:
            pres.Saved = True
            pres.Close()
    finally:
        app.Quit()

    rows.sort(key=lambda r: (r[3][1], r[3][0]))
    print(f"{'arm':<16}{'PPT lines':>10}{'engine':>8}{'BoundHeight':>13}   verdict")
    base = next((h for lbl, _, h, _ in rows if lbl.startswith("A_")), None)
    for label, n, h, (left, top) in rows:
        eng = mine.get((round(left, 1), round(top, 1)))
        tall = "" if base is None else f"  ({h / base:.2f}x the one-line box)"
        flag = "" if eng == n else "   <-- DISAGREES"
        print(f"{label:<16}{n:>10}{eng if eng is not None else '-':>8}{h:>13.2f}{tall}{flag}")


if __name__ == "__main__":
    main()
