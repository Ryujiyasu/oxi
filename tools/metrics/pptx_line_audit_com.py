# -*- coding: utf-8 -*-
"""Where PowerPoint breaks each paragraph, against where the engine breaks it.

Every pptx measurement so far has gone through the truth PDF, and the PDF is a
poor oracle for geometry: PowerPoint restates a line there as a `Tf` size that
is not the declared one, plus a per-run `Tc`, plus sparse integer `TJ`, so the
effective advance wobbles +-0.9% with the declared size
(`read_pptx_drawgrid_com.py`). The application itself has no such problem --
`TextRange.Paragraphs(i).Lines.Count` is PowerPoint's own answer to the one
question that matters most, because a paragraph broken into a different number
of lines is not a sub-pixel disagreement, it is a different layout.

So this asks PowerPoint directly and compares against the engine's own
`--dump-layout`:

    line count per paragraph   the BREAK -- categorical, cause-attributable
    line left edge             the ALIGNMENT, as a secondary number

It is the pptx analogue of the docx pagination gate: a signal, not an outcome.

★Two things must not overlap this: the renderer (`pptx_render_not_parallel_safe`)
and any other PowerPoint COM session (`pptx_com_render_must_not_overlap`). The
dump is taken first, then COM is opened, never both at once.

    python tools/metrics/pptx_line_audit_com.py 34 [48 33 ...]
    python tools/metrics/pptx_line_audit_com.py --all
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"

# How close two shapes must sit to be the same shape. The dump and COM both
# answer in points from the slide's top left, so this only absorbs the rounding
# each side does on the way out.
NEAR = 0.75


def engine_dump(src: Path) -> dict:
    """The engine's layout for `src`, via `--dump-layout`."""
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "layout.json"
        subprocess.run(
            [str(EXE), str(src), str(Path(td) / "slide"), "150", f"--dump-layout={out}"],
            capture_output=True, timeout=3600, check=False)
        if not out.exists():
            return {}
        return json.loads(out.read_text(encoding="utf-8"))


def engine_paras(dump: dict) -> dict[int, list]:
    """Per slide: (x, y, w, [(line_count, x_offsets, text), ...]) per text shape."""
    out: dict[int, list] = {}
    for si, slide in enumerate(dump.get("slides", []), start=1):
        shapes = []
        for sh in slide.get("shapes", []):
            content = sh.get("content") or {}
            paras = content.get("paragraphs")
            if not paras:
                continue
            rows = []
            for p in paras:
                text = "".join(r.get("text", "") for r in p.get("runs", []))
                rows.append((len(p.get("line_x_offsets") or []),
                             p.get("line_x_offsets") or [], text))
            shapes.append({"x": sh["x"], "y": sh["y"], "w": sh["w"], "paras": rows})
        out[si] = shapes
    return out


def com_shapes(shape, acc: list) -> None:
    """Flatten groups the way the parser does, keeping absolute geometry."""
    # 6 = msoGroup
    if shape.Type == 6:
        for i in range(1, shape.GroupItems.Count + 1):
            com_shapes(shape.GroupItems(i), acc)
        return
    try:
        if not shape.HasTextFrame or not shape.TextFrame.HasText:
            return
    except Exception:
        return
    try:
        tr = shape.TextFrame.TextRange
        rows = []
        for pi in range(1, tr.Paragraphs().Count + 1):
            para = tr.Paragraphs(pi)
            # ★`para.Lines()` returns the whole range; the INDEX goes to the
            # METHOD, not to what it returned. `para.Lines()(j)` calls a
            # TextRange and raises "does not support collection" -- which the
            # guard below swallowed, so a deck full of text audited as
            # "0 shapes" and read like a clean pass.
            n = para.Lines().Count
            rows.append((n,
                         [round(para.Lines(j).BoundLeft - shape.Left, 2)
                          for j in range(1, n + 1)],
                         para.Text.rstrip("\r")))
        acc.append({"x": shape.Left, "y": shape.Top, "w": shape.Width, "paras": rows})
    except Exception as e:
        # Loud, not silent: a shape this cannot read is a hole in the audit's
        # coverage, and a coverage hole that prints nothing reads as a pass.
        print(f"      shape {shape.Name[:28]!r} refused: {str(e)[:60]}", flush=True)
        return


def audit(doc: int) -> dict | None:
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    item = next((i for i in manifest if i["idx"] == doc), None)
    if item is None:
        return None
    src = ROOT / "pptx" / item["local"]
    if not src.exists():
        return None
    eng = engine_paras(engine_dump(src))
    if not eng:
        print(f"{doc:02d}: the engine produced no layout", flush=True)
        return None

    app = win32com.client.Dispatch("PowerPoint.Application")
    got = {"paras": 0, "same": 0, "diff": 0, "shapes": 0, "unmatched": 0}
    worst: list = []
    try:
        pres = app.Presentations.Open(str(src.resolve()), WithWindow=False)
        try:
            for si in range(1, pres.Slides.Count + 1):
                acc: list = []
                slide = pres.Slides(si)
                for i in range(1, slide.Shapes.Count + 1):
                    com_shapes(slide.Shapes(i), acc)
                mine = eng.get(si, [])
                for c in acc:
                    got["shapes"] += 1
                    m = next((e for e in mine
                              if abs(e["x"] - c["x"]) < NEAR
                              and abs(e["y"] - c["y"]) < NEAR
                              and abs(e["w"] - c["w"]) < NEAR), None)
                    if m is None:
                        got["unmatched"] += 1
                        continue
                    for (cn, cx, ctext), (en, ex, etext) in zip(c["paras"], m["paras"]):
                        got["paras"] += 1
                        if cn == en:
                            got["same"] += 1
                        else:
                            got["diff"] += 1
                            worst.append((si, cn, en, ctext[:44]))
        finally:
            pres.Saved = True
            pres.Close()
    finally:
        app.Quit()
    rate = 100.0 * got["same"] / got["paras"] if got["paras"] else 0.0
    print(f"{doc:02d}: {got['paras']:5} paragraphs  {rate:6.2f}% break agreement  "
          f"({got['diff']} differ)  shapes {got['shapes']} "
          f"({got['unmatched']} unmatched)", flush=True)
    for si, cn, en, t in worst[:6]:
        print(f"      s{si:<3} PowerPoint {cn} lines, engine {en}  {t!r}", flush=True)
    return got


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("docs", nargs="*", type=int)
    ap.add_argument("--all", action="store_true")
    args = ap.parse_args()
    docs = args.docs
    if args.all:
        manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
        docs = [i["idx"] for i in manifest]
    if not docs:
        sys.exit("name at least one deck, or pass --all")
    tot = {"paras": 0, "same": 0, "diff": 0}
    for d in docs:
        g = audit(d)
        if g:
            for k in tot:
                tot[k] += g[k]
    if tot["paras"]:
        print(f"\n{tot['paras']} paragraphs over {len(docs)} decks: "
              f"{100.0 * tot['same'] / tot['paras']:.2f}% break agreement, "
              f"{tot['diff']} differ")


if __name__ == "__main__":
    main()
