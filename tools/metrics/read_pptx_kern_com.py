# -*- coding: utf-8 -*-
"""Read the kern sweep: PowerPoint's own character steps against the design.

Per arm, the step PowerPoint puts between each pair is compared with the first
glyph's design advance out of the font file. The difference is the kern, in
points and in em, so the arms can be read against each other:

    a `pair` arm negative and the `flat` control zero   PowerPoint kerns
    `kern9600` zero and `kern1200` negative at 40pt     `@kern` is a minimum size
    `pair20` half of `pair40`                           the amount scales

★Must not overlap the renderer or another COM session.

    python tools/metrics/read_pptx_kern_com.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import win32com.client
from fontTools.ttLib import TTFont

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
from pptx_gdi_face_audit import registry_file  # noqa: E402
from gen_pptx_kern import ARMS, FACE  # noqa: E402

DECK = REPO / "pipeline_data" / "pptx_probes" / "kern" / "kern.pptx"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def design_advances() -> tuple[dict[str, float], int]:
    path = registry_file(FACE)
    if path is None:
        sys.exit(f"{FACE} is not installed, so there is nothing to compare against")
    font = TTFont(str(path))
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    adv = {}
    for ch in set("".join(text for _, text, _, _ in ARMS)):
        g = cmap.get(ord(ch))
        if g is not None:
            adv[ch] = hmtx[g][0] / upm
    return adv, upm


def main() -> None:
    if not DECK.exists():
        sys.exit(f"{DECK} is not there -- run gen_pptx_kern.py first")
    adv, upm = design_advances()

    app = win32com.client.Dispatch("PowerPoint.Application")
    seen: dict[str, list[float]] = {}
    try:
        pres = app.Presentations.Open(str(DECK.resolve()), WithWindow=False)
        try:
            slide = pres.Slides(1)
            labels, arms = [], []
            wanted = {text for _, text, _, _ in ARMS}
            for i in range(1, slide.Shapes.Count + 1):
                sh = slide.Shapes(i)
                if not sh.HasTextFrame or not sh.TextFrame.HasText:
                    continue
                text = sh.TextFrame.TextRange.Text.strip()
                if text in wanted:
                    arms.append((sh, text))
                else:
                    labels.append((sh.Left, sh.Top, text))
            for sh, text in arms:
                tr = sh.TextFrame.TextRange
                above = [(t, n) for (x, t, n) in labels
                         if abs(x - sh.Left) < 1.0 and t < sh.Top]
                label = max(above)[1] if above else f"x={sh.Left:.0f}"
                lefts = [tr.Characters(k, 1).BoundLeft
                         for k in range(1, len(text) + 1)]
                seen[label] = [round(b - a, 4) for a, b in zip(lefts, lefts[1:])]
        finally:
            pres.Saved = True
            pres.Close()
    finally:
        app.Quit()

    print(f"steps PowerPoint puts between characters, against {FACE}'s design\n")
    print(f"{'arm':<10}{'pair':<8}{'step':>9}{'design':>9}{'kern':>9}{'kern em':>10}")
    for label, text, size, kern in ARMS:
        steps = seen.get(label)
        if not steps:
            print(f"{label:<10}(not read)")
            continue
        for k, step in enumerate(steps):
            pair = text[k:k + 2]
            want = adv.get(text[k])
            if want is None:
                continue
            want_pt = want * size
            print(f"{label:<10}{pair!r:<8}{step:>9.3f}{want_pt:>9.3f}"
                  f"{step - want_pt:>+9.3f}{(step - want_pt) / size:>+10.4f}")
        print()


if __name__ == "__main__":
    main()
