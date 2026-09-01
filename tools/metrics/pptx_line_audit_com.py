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

★FIRST RESULT (2026-09-02): **2428 paragraphs over 16 blind decks, 100.00%
agreement, 0 differ** -- including 48 and 33, the two decks S-MUDRAW cost
ground. So the remaining pptx gap is entirely WITHIN the line; there are no
break errors left to chase on this corpus.

★And the NEGATIVE CONTROL, because a 100% from an instrument that cannot fail
is worth nothing: with `OXI_MASTERUNIT_DISABLE=1` the same run reports
`09 s2  PowerPoint 3 lines, engine 4` on the template-instructions paragraph.
The tool can see a break disagreement, so not seeing one is evidence. Run that
control again whenever this reports a clean sweep.

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
            # ★A turned shape's lines run along an axis the dump does not turn:
            # the engine answers in the shape's own frame while `BoundLeft` is
            # slide space, so comparing them measures the rotation. d04 and d21
            # both put their worst offsets on s34, whose 'HIGH VALUE 1' sits at
            # +90 -- the same shape the editor sweep skips, for the same reason.
            if sh.get("rotation"):
                continue
            rows = []
            for p in paras:
                text = "".join(r.get("text", "") for r in p.get("runs", []))
                rows.append((len(p.get("line_x_offsets") or []),
                             p.get("line_x_offsets") or [], text,
                             p.get("line_baselines") or []))
            # ★The dump's line offsets are measured from the TEXT AREA, while
            # PowerPoint's `BoundLeft` is measured from the slide. The engine
            # STATES where its text area starts (`text_left`) rather than this
            # rebuilding the inset rules -- which matters, because an audit that
            # recomputes the same insets can never catch the insets being wrong.
            #
            # Both biases showed themselves before the field existed: deck 21's
            # median came back exactly +7.20pt (`l_ins`), and its three ellipses
            # read +26.4pt, which is an ellipse's own text inset on a 180.38pt
            # box to the hundredth.
            shapes.append({"x": sh["x"], "y": sh["y"], "w": sh["w"],
                           "text_left": sh.get("text_left", sh["x"]),
                           "paras": rows})
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
        # Skipped on BOTH sides, or a shape the engine side drops reads as
        # "unmatched" -- a coverage hole dressed up as a mismatch.
        if abs(shape.Rotation) > 1e-6:
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
                         para.Text.rstrip("\r"),
                         [round(para.Lines(j).BoundTop, 2)
                          for j in range(1, n + 1)]))
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
    got = {"paras": 0, "same": 0, "diff": 0, "shapes": 0, "unmatched": 0,
           "para_count": 0}
    worst: list = []
    offs: list[float] = []
    far: list = []
    vadv: list[float] = []
    vfar: list = []
    pfar: list = []
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
                    near = [e for e in mine
                            if abs(e["x"] - c["x"]) < NEAR
                            and abs(e["y"] - c["y"]) < NEAR
                            and abs(e["w"] - c["w"]) < NEAR]
                    # ★Identity before measurement. Deck 33 s7 stacks THREE
                    # groups at x=1146.19 w=155.82 -- 'Upcoming', 'Done' and
                    # 'Finalizing' -- so taking the first geometric match paired
                    # one label's box with another's text and reported 47.15pt
                    # and 28.02pt of "defect" that was entirely this. Geometry
                    # narrows the candidates; the TEXT decides between them, and
                    # a tie is counted as unmatched rather than guessed at.
                    want = [p[2] for p in c["paras"]]
                    exact = [e for e in near if [p[2] for p in e["paras"]] == want]
                    m = exact[0] if len(exact) == 1 else (
                        near[0] if len(near) == 1 else None)
                    if m is None:
                        got["unmatched"] += 1
                        continue
                    # ★Two paragraph lists of different LENGTH must not be
                    # zipped. Deck 47 s2 has 3 paragraphs in PowerPoint and 4 in
                    # the engine, so from the first mismatch on, every
                    # comparison was between unrelated paragraphs -- and the
                    # deck's 6pt `spcBef` came out as a phantom "+6.00pt line
                    # advance error" on 39 lines. The count disagreement is the
                    # real finding; it is reported as itself.
                    # ★A TRAILING EMPTY paragraph is in the file and in the
                    # engine, but `TextRange.Paragraphs()` does not enumerate
                    # it. Deck 47 s2 really does hold four `<a:p>` with the
                    # last one empty; COM says three. Dropping it here is
                    # normalising a representational difference, not hiding a
                    # defect -- the engine is right to keep it, and it is what
                    # made twelve shapes look mismatched.
                    mine_paras = m["paras"]
                    if (len(mine_paras) == len(c["paras"]) + 1
                            and not mine_paras[-1][2].strip()):
                        mine_paras = mine_paras[:-1]
                    m = dict(m, paras=mine_paras)
                    if len(c["paras"]) != len(m["paras"]):
                        got["para_count"] += 1
                        pfar.append((si, len(c["paras"]), len(m["paras"]),
                                     c["paras"][0][2][:38] if c["paras"] else ""))
                        continue
                    for (cn, cx, ctext, cy), (en, ex, etext, ey) in zip(
                            c["paras"], m["paras"]):
                        got["paras"] += 1
                        if cn == en:
                            got["same"] += 1
                            # Where PowerPoint put the line's left edge against
                            # where the engine put it.
                            #
                            # ★`BoundLeft` reads as a PEN position, not an ink
                            # one: the median of this difference comes back
                            # -0.00pt on d34 and +0.00pt on d48, which a left
                            # side bearing would not do. Measured, not assumed
                            # -- the name says "Bound", so the median is
                            # reported every run as the check on that reading.
                            # The SPREAD is the signal either way, since a
                            # constant bias would cancel out of it.
                            # PowerPoint's absolute line left, against the
                            # engine's own text-area origin plus its offset.
                            # ★The line-to-line ADVANCE, which is what
                            # accumulates. `BoundTop` is an ink top and the
                            # engine's is a baseline, so they differ by an
                            # ascent -- a per-face constant that cancels out of
                            # consecutive differences. Deck 47 is the specimen:
                            # its text doubles vertically down a block, which no
                            # horizontal number can see.
                            for k in range(1, min(len(cy), len(ey))):
                                step = (cy[k] - cy[k - 1]) - (ey[k] - ey[k - 1])
                                vadv.append(step)
                                if abs(step) > 0.5:
                                    vfar.append((si, round(step, 2), ctext[:38]))
                            for a, b in zip(cx, ex):
                                d = (a + c["x"]) - (m["text_left"] + b)
                                offs.append(d)
                                if abs(d) > 3.0:
                                    far.append((si, round(d, 2), ctext[:40]))
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
    if offs:
        srt = sorted(offs)
        med = srt[len(srt) // 2]
        spread = sorted(abs(v - med) for v in offs)
        got["lines"] = len(offs)
        got["p95"] = spread[int(0.95 * (len(spread) - 1))]
        print(f"      line left: median {med:+.2f}pt (0 = BoundLeft is a pen), "
              f"spread |x-median| p50 {spread[len(spread) // 2]:.2f}pt "
              f"p95 {got['p95']:.2f}pt, {len(far)} over 3pt", flush=True)
        for si, d, t in sorted(far, key=lambda r: -abs(r[1]))[:4]:
            print(f"        s{si:<3} {d:+7.2f}pt  {t!r}", flush=True)
    if got["para_count"]:
        print(f"      shapes whose PARAGRAPH COUNT disagrees: {got['para_count']} "
              f"(not compared further -- the lists do not line up)", flush=True)
        for si, cn, en, txt in pfar[:4]:
            print(f"        s{si:<3} PowerPoint {cn} paragraphs, engine {en}  {txt!r}",
                  flush=True)
    if vadv:
        srt = sorted(abs(v) for v in vadv)
        got["vadv_p95"] = srt[int(0.95 * (len(srt) - 1))]
        print(f"      line advance: {len(vadv)} steps, |err| p50 "
              f"{srt[len(srt) // 2]:.2f}pt p95 {got['vadv_p95']:.2f}pt, "
              f"{len(vfar)} over 0.5pt", flush=True)
        for si, d, t in sorted(vfar, key=lambda r: -abs(r[1]))[:4]:
            print(f"        s{si:<3} {d:+7.2f}pt  {t!r}", flush=True)
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
