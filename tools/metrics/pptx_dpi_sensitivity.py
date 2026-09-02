# -*- coding: utf-8 -*-
"""How much of the corpus's wrapping depends on the DPI it is measured at.

A break the master-unit law can answer is scale-free by construction, so every
paragraph that moves between two DPIs is one whose break fell through to GDI's
device-pixel extent -- and those are the paragraphs where two callers passing
two different scales disagree about the layout.

That is not hypothetical. `dump_layout_json_gdi` measured at 1.0 while the
renderer draws at `dpi * supersample / 72`, so the break audit was reading a
layout the picture never had (blind 31 s21: one line dumped, two lines drawn).
That one is fixed. **`compute_shape_anchor_off` still measures a text block's
height at 1.0**, so a scale-sensitive paragraph inside a centred or
bottom-anchored shape is still anchored against a block height that is not the
one drawn -- and unlike the dump, that shows in the picture.

★What it says today (2026-09-02, 72 vs 150 DPI, dev + blind):

    48365 paragraphs over 114 decks   0 break differently, 8439 shift in-line

So the categorical exposure is now nil: with S-FDBREAK reading the part's own
design table by default, no paragraph in the corpus changes its LINE COUNT with
the resolution any more (before it, blind 31 had two). The in-line shifts are
the device-pixel rounding of positions and are expected. Re-run this whenever a
face stops answering an advance -- that is the condition that re-opens it.

    python tools/metrics/pptx_dpi_sensitivity.py            # dev + blind
    python tools/metrics/pptx_dpi_sensitivity.py 31 72 150  # one deck, two DPIs
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
from pptx_dump_ab import deck_paths, paragraphs, wait_for_powerpoint_to_exit  # noqa: E402

EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def dump_at(src: Path, dpi: int) -> dict | None:
    """The engine's layout for `src`, measured as it would be drawn at `dpi`."""
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "layout.json"
        subprocess.run(
            [str(EXE), str(src), str(Path(td) / "slide"), str(dpi),
             f"--dump-layout={out}"],
            capture_output=True, env=dict(os.environ), timeout=3600, check=False)
        if not out.exists():
            return None
        return json.loads(out.read_text(encoding="utf-8"))


def main() -> None:
    spec = sys.argv[1] if len(sys.argv) > 1 else "all"
    lo, hi = (int(sys.argv[2]), int(sys.argv[3])) if len(sys.argv) > 3 else (72, 150)
    decks = deck_paths(spec)
    if not decks:
        sys.exit("no decks selected")
    total = differ = shifted = 0
    hits: list[str] = []
    for name, src in decks:
        wait_for_powerpoint_to_exit()
        a, b = dump_at(src, lo), dump_at(src, hi)
        if a is None or b is None:
            print(f"{name}: render failed", flush=True)
            continue
        pa, pb = paragraphs(a), paragraphs(b)
        shared = pa.keys() & pb.keys()
        total += len(shared)
        moved = [(k, pa[k], pb[k]) for k in shared if pa[k][0] != pb[k][0]]
        slid = [k for k in shared if pa[k][0] == pb[k][0] and pa[k][1] != pb[k][1]]
        differ += len(moved)
        shifted += len(slid)
        if moved:
            hits.append(f"{name}({len(moved)})")
            print(f"{name}: {len(moved)} paragraphs break differently at {lo} vs {hi} "
                  f"DPI, {len(slid)} shift, of {len(shared)}", flush=True)
            for k, x, y in moved[:4]:
                print(f"      s{k[0]:<3} {x[0]} -> {y[0]} lines  {x[2][:44]!r}", flush=True)
    print(f"\n{total} paragraphs over {len(decks)} decks: {differ} break differently "
          f"between {lo} and {hi} DPI, {shifted} shift within the line")
    if hits:
        print("decks: " + " ".join(hits))


if __name__ == "__main__":
    main()
