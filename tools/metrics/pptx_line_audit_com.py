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
    line width                 the ADVANCES, which is what is left once the
                               breaks agree (`Lines(j).BoundWidth`)

It is the pptx analogue of the docx pagination gate: a signal, not an outcome.

★FIRST RESULT (2026-09-02): 2428 paragraphs over 16 blind decks, 100.00%
agreement, 0 differ -- including 48 and 33, the two decks S-MUDRAW cost ground.

★THE LEDGER (2026-09-02, the whole blind set, after the dump was fixed to
measure at the scale the picture is drawn at): **7620 paragraphs over 48 decks,
4 disagreements in 2 decks**. All four have since been read:

    44 s25, 44 s36, d06 s36, d19 s37   the paragraph ends in `<a:br/>`, and
        `Lines.Count` does not count the empty line it opens even though
        PowerPoint reserves its height (`gen_pptx_trailbr.py`). The ENGINE was
        right; this tool now compares them on the same footing.
    47 s2, 47 s25   this machine's GDI serves Caladea 3.36% narrow
        (`pptx_gdi_face_audit.py`), so the engine fits a line PowerPoint does
        not. Environmental, not a defect (`gdi_font_view_can_be_corrupt`).

So the remaining pptx gap is entirely WITHIN the line on this corpus. Two of
this tool's own blind spots were closed on the way: turned shapes were dropped
whole (2.8% of the corpus's text shapes) and are now compared for their breaks
AND their widths, and the trailing break above.

★A turned shape's WIDTH is comparable even though its left edge is not: the
length of a line does not depend on which way the shape faces, only on which of
the box's two axes reports it. 169 of the corpus's 252 turned text shapes sit at
exactly 90 or 270 degrees (`BoundHeight` carries the width there) and most of
the rest at 180 (`BoundWidth`, as usual); anything off a right angle mixes the
axes and is left out. On d39 that takes the width comparison from 63 lines to
103, and the 40 it adds agree (median +0.03pt, 0 over 2%).

★THE WIDTH, added once the breaks were clean (2026-09-02). PowerPoint's
`BoundWidth` is an INK box and the engine's `line_w` is a PEN advance, so the
pair is only comparable after three confounders are taken out, each of which
first appeared as a fat defect:

    a MARKER paragraph      PowerPoint's box takes in the bullet, the engine's
                            width is the text alone -- d09's '1.' read +39.6%
    a TRAILING SPACE        the engine trims, `BoundWidth` runs past it --
                            d09 s7's 'BIG ' read +12.94pt, and a 52pt Playfair
                            space is 13pt
    a SHORT line            a side bearing is a large share of 20pt and none of
                            200pt, so short lines are listed but kept out of the
                            fit, which would otherwise read the intercept as a
                            slope

What is left is reported as a fit: `ink - pen = a * width + b`, with `a` in per
mille (the advance error, which grows with the line) and `b` in points (the
bearing, which does not). Cleaned up, deck 9 reads median +0.03pt and deck 34
+0.00pt -- the engine's advances agree with PowerPoint's.

★And its own control: pointed at d32, whose Bebas Neue was already on record as
mismeasured, the top line is s1's 'we help' at **+62.02pt (+10.9%)** -- 223pt,
`b="1"`, in a family with no bold. So the number can find a real width defect,
which is what makes the near-zero medians elsewhere evidence.

★What the box IS, settled by `gen_pptx_boundwidth.py` (one word, one face,
one size, one property changed per arm):

    alone in its shape          box - pen = -1.27pt at 14pt, -2.42 at 28
                                (an ink box: the end glyphs' bearings)
    autofit / nowrap / centre /
    right / insets              all identical to the hundredth -- 39.97pt
    a SECOND paragraph present  box - pen = **+2.60pt**, first or second
                                paragraph, bold or not

The step between those two is 3.87pt and Arial's space at 14pt is 3.889pt: in a
multi-paragraph shape the box takes in the PARAGRAPH MARK. It scales, too --
the 28pt arm steps 7.74pt against a 7.778pt space -- so the engine now states
that space per paragraph (`space_pt`) and `--mark-correct` will subtract it.
That stays OPT-IN: a real deck does not apply the mark as predictably as the
probe does (deck 9's single-line 'THREATS', inside a multi-paragraph
placeholder, reads -3.31pt once corrected, which says its box never had one), so
the gate keeps the population whose null it can account for. That was the whole of
the "19% of lines sit 2pt out" class -- including d09 s9's 'Yellow' and d32 s1's
+62pt, neither of which is a defect. With single-paragraph shapes only:

    9 / d32 / 34    232 lines, median -0.00pt, slope under 0.7 per mille,
                    **0 lines over 2%**

★Its POSITIVE CONTROL is `OXI_ADVEXACT_DISABLE=1`, and it took some finding:

    deck 9, default              median -0.00pt   slope  +0.05 per mille   0 over 2%
    deck 9, ADVEXACT off         median -0.24pt   slope **-35.22** per mille  3 over 2%
                                 worst line -90.40pt (-20.9%)

Five flags that steer FACE and ADVANCE decisions -- `OXI_CLOUDADV_DISABLE`,
`OXI_SLOTFACE_DISABLE`, `OXI_FDSYNTH_DISABLE`, `OXI_BOLDADV_DISABLE`,
`OXI_HMTXSTYLE_DISABLE` -- leave the figure identical to the hundredth, which is
itself worth knowing: `line_w` comes from `per_run` / `runtime_width_px` and the
break test from `advance_em`, so the two gates in this file watch DIFFERENT
chains and a change to one need not show in the other.

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

# Whether shapes that carry a rotation are audited at all (breaks only).
ROTATED = True

# When set, every compared line's width pair is collected here and written out.
# Four examples in a console are a hypothesis; a distribution is a finding.
WIDTH_ROWS: list | None = None

# Whether to compare multi-paragraph shapes by subtracting the paragraph mark.
MARK_CORRECT = False

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"

# How close two shapes must sit to be the same shape. The dump and COM both
# answer in points from the slide's top left, so this only absorbs the rounding
# each side does on the way out.
NEAR = 0.75


def wait_for_powerpoint_to_exit(limit: float = 60.0) -> None:
    """Block until no POWERPNT.EXE is left running.

    ★`app.Quit()` returns before the process is gone, and a renderer started in
    that window resolves embedded fonts against a PowerPoint that still holds
    them (`pptx_com_render_must_not_overlap`). This tool audited one deck at a
    time for most of its life and never noticed; run twelve decks in a row and
    d44 reports 16 break disagreements that a single-deck run does not have --
    the same deck, the same binary, 88.81% against 100.00%.
    """
    import time
    deadline = time.time() + limit
    while time.time() < deadline:
        r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
                           capture_output=True, text=True, check=False)
        if "POWERPNT" not in (r.stdout or ""):
            return
        time.sleep(0.5)


def engine_dump(src: Path) -> dict:
    """The engine's layout for `src`, via `--dump-layout`."""
    wait_for_powerpoint_to_exit()
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
            #
            # ★But that is an argument about GEOMETRY, and it was costing the
            # BREAK its evidence too: 390 of the corpus's 14002 text shapes are
            # turned (2.8%, and 29% of d39 -- `pptx_audit_blindspot.py`), so a
            # clean sweep was a claim about the other 97%. A line COUNT does not
            # care which way the shape faces. So a turned shape stays in, with
            # its geometry excluded and only its breaks compared.
            turned = bool(sh.get("rotation"))
            if turned and not ROTATED:
                continue
            rows = []
            for p in paras:
                text = "".join(r.get("text", "") for r in p.get("runs", []))
                # ★A paragraph with a MARKER is left out of the width
                # comparison: PowerPoint's line box takes in the bullet or the
                # number, the engine's `line_w` is the text alone, and the
                # difference is the marker rather than any advance. d09's '1.'
                # read +8.94pt (+39.6%) that way, which is a bullet, not a
                # defect. The line count and the left edge are unaffected.
                marked = bool(p.get("marker"))
                # ★A line whose text ends in a space is left out of the width
                # comparison as well: the engine measures `line_w` on the line
                # TRIMMED and PowerPoint's `BoundWidth` runs to the pen after
                # the space, so the pair differs by a space and not by anything
                # either side got wrong. d09 s7's 'BIG ' read +12.94pt that way
                # -- a 52pt Playfair space is 13pt.
                widths = [] if marked else (p.get("line_widths") or [])
                texts = p.get("line_texts") or []
                if widths and len(texts) == len(widths):
                    widths = [w if not (t.endswith(" ") or t.endswith("\n")) else 0.0
                              for w, t in zip(widths, texts)]
                rows.append((len(p.get("line_x_offsets") or []),
                             p.get("line_x_offsets") or [], text,
                             p.get("line_baselines") or [], widths,
                             p.get("measured_family") or "",
                             float(p.get("space_pt") or 0.0)))
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
                           "turned": turned,
                           "rot": float(sh.get("rotation") or 0.0),
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
        # "unmatched" -- a coverage hole dressed up as a mismatch. With
        # `ROTATED` it is kept on both sides instead, and only its GEOMETRY is
        # left out of the comparison.
        turned = abs(shape.Rotation) > 1e-6
        if turned and not ROTATED:
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
                          for j in range(1, n + 1)],
                         # ★What PowerPoint set the line AT. `BoundWidth` is an
                         # ink box and the engine's number is a pen advance, so
                         # the two differ by the first and last glyph's side
                         # bearings -- a constant that cancels out of the SPREAD
                         # and out of the slope against line length, which is
                         # where an advance error lives.
                         [round(para.Lines(j).BoundWidth, 2)
                          for j in range(1, n + 1)],
                         # ★And the box's other axis, because a quarter turn
                         # swaps them: 169 of the corpus's 252 turned text
                         # shapes sit at exactly 90 or 270 degrees, where the
                         # line's WIDTH is what `BoundHeight` reports.
                         [round(para.Lines(j).BoundHeight, 2)
                          for j in range(1, n + 1)],
                         # What the RUN asks for. Beside the family the engine
                         # actually measured, this says whether a width
                         # disagreement is about advances or about which face
                         # answered -- the two need different fixes.
                         (para.Font.Name or "")))
        acc.append({"x": shape.Left, "y": shape.Top, "w": shape.Width,
                    "turned": turned, "rot": float(shape.Rotation or 0.0),
                    "paras": rows})
    except Exception as e:
        # Loud, not silent: a shape this cannot read is a hole in the audit's
        # coverage, and a coverage hole that prints nothing reads as a pass.
        print(f"      shape {shape.Name[:28]!r} refused: {str(e)[:60]}", flush=True)
        return


def deck_path(doc: str) -> Path | None:
    """A blind deck by index, or a dev deck by its `dNN` name.

    Dev decks need no truth PDF for this audit -- PowerPoint is the oracle --
    so the corpus can grow without paying for an export per deck.
    """
    if str(doc).lower().startswith("d"):
        hit = sorted((ROOT / "dev" / "pptx").glob(f"{doc}__*.pptx"))
        return hit[0] if hit else None
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    item = next((i for i in manifest if i["idx"] == int(doc)), None)
    if item is None:
        return None
    src = ROOT / "pptx" / item["local"]
    return src if src.exists() else None


def audit(doc) -> dict | None:
    src = deck_path(doc)
    if src is None:
        print(f"{doc}: no such deck", flush=True)
        return None
    eng = engine_paras(engine_dump(src))
    if not eng:
        print(f"{doc:02d}: the engine produced no layout", flush=True)
        return None

    app = win32com.client.Dispatch("PowerPoint.Application")
    got = {"paras": 0, "same": 0, "diff": 0, "shapes": 0, "unmatched": 0,
           "para_count": 0, "turned_paras": 0, "trailing_br": 0}
    worst: list = []
    offs: list[float] = []
    far: list = []
    vadv: list[float] = []
    vfar: list = []
    pfar: list = []
    # (engine's pen width, PowerPoint's ink width minus it) per line.
    wide: list[tuple[float, float]] = []
    wfar: list = []
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
                    # A turned shape is compared for its BREAKS only: its line
                    # boxes are stated in two different frames, and rotating one
                    # into the other would make this audit depend on arithmetic
                    # it is supposed to be checking.
                    geom = not (c.get("turned") or m.get("turned"))
                    # Which of PowerPoint's two box axes carries the line's
                    # WIDTH at this shape's turn. A quarter turn swaps them, a
                    # half turn keeps them, and anything else mixes the two and
                    # is left out.
                    rot = (c.get("rot", 0.0) or 0.0) % 360.0
                    quart = round(rot / 90.0) * 90 % 360
                    axis = (None if abs(rot - quart) > 0.5
                            else (0 if quart in (0, 180) else 1))
                    for (cn, cx, ctext, cy, cw, chh, cface),                             (en, ex, etext, ey, ew, eface, espace) in zip(
                            c["paras"], m["paras"]):
                        got["paras"] += 1
                        if not geom:
                            got["turned_paras"] += 1
                        # ★The oracle's own blind spot. `Lines.Count` does not
                        # count the empty line a TRAILING `<a:br/>` opens, but
                        # PowerPoint reserves its height: for 'abc' + br it says
                        # one line and reports `BoundHeight` 43.20 against a
                        # one-line box's 21.60, and for two trailing breaks it
                        # says two and reports 64.80 (`gen_pptx_trailbr.py`,
                        # seven arms in both wrap modes). The four paragraphs in
                        # dev + blind that end in a break -- 44 s25/s36, d06 s36,
                        # d19 s37 -- were the whole of this audit's 0.05%, and
                        # the engine was right in all four.
                        if etext.endswith("\n") and en > 1:
                            en -= 1
                            got["trailing_br"] += 1
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
                            #
                            # ★And the FIRST step is skipped, because
                            # `Lines(1).BoundTop` bounds the paragraph's
                            # space-before as well as its text while lines 2+
                            # are ink tops. Deck 47 s2 p3: PowerPoint steps
                            # [22.56, 16.56] against the engine's
                            # [16.56, 16.56], and 22.56 - 16.56 is exactly the
                            # 6pt `spcBef` that paragraph declares. Every
                            # "+6.00pt advance error" in this deck was that.
                            if geom:
                                for k in range(2, min(len(cy), len(ey))):
                                    step = (cy[k] - cy[k - 1]) - (ey[k] - ey[k - 1])
                                    vadv.append(step)
                                    if abs(step) > 0.5:
                                        vfar.append((si, round(step, 2), ctext[:38]))
                                for a, b in zip(cx, ex):
                                    d = (a + c["x"]) - (m["text_left"] + b)
                                    offs.append(d)
                                    if abs(d) > 3.0:
                                        far.append((si, round(d, 2), ctext[:40]))
                            if axis is not None:
                                # ★The width is compared for a TURNED shape too,
                                # unlike the two numbers above: a line's length
                                # does not depend on which way the shape faces,
                                # only on which of the box's axes reports it,
                                # and `axis` has just decided that. The left edge
                                # and the line advance stay out because they are
                                # stated in two different frames.
                                # The line's WIDTH, which is the only number
                                # here that grows with the error it carries: a
                                # side bearing is the same on a 20pt line and a
                                # 200pt one, an advance that is 1% wrong is not.
                                # Short lines are dropped from the FIT, not
                                # from the listing: a side bearing is a large
                                # share of a 20pt line and none of a 200pt one,
                                # so leaving them in lets the intercept masquerade
                                # as a slope.
                                #
                                # ★And only the LAST line of a paragraph is
                                # compared at all. Every earlier line broke AT a
                                # space, the wrap trims it (`WrapOpts
                                # .trim_trailing_space`) and PowerPoint's box
                                # does not -- so those pairs differ by one space
                                # and nothing else. It shows as the same
                                # ABSOLUTE delta repeated across unrelated texts
                                # in a deck (+11.19pt on 'Add a' and on 'Icons',
                                # +4.19pt on 'Add a' and on 'Black'), which is a
                                # constant, and a constant is not an advance
                                # error. The trimmed space is not in
                                # `line_texts` either, so it cannot be filtered
                                # by looking at them.
                                # ★And only in a shape that holds ONE
                                # paragraph. `gen_pptx_boundwidth.py` settles
                                # why: the same word in the same face and size
                                # measures `box - pen = -1.27pt` alone in its
                                # shape -- an ink box, narrower than the pen by
                                # the end bearings -- and `+2.60pt` as soon as
                                # the shape holds a second paragraph, first or
                                # second, bold or not. The step is 3.87pt, and
                                # Arial's space at 14pt is 3.889pt: the box
                                # takes in the paragraph mark. That is the whole
                                # of the "19% of lines are 2pt out" class, d09
                                # s9's 'Yellow' included, and none of it is the
                                # engine.
                                # ★A shape with more than one paragraph has
                                # the paragraph MARK inside its box, and the
                                # mark is exactly one space wide. The engine
                                # states that space (`space_pt`), so those
                                # shapes are compared with it taken off rather
                                # than dropped -- which is most of the corpus.
                                # ...but only in the arrangement the probe
                                # actually measured: a SINGLE-LINE paragraph in
                                # a multi-paragraph shape. Applying it to a
                                # WRAPPED paragraph's last line turned deck 9's
                                # clean sheet into 19 lines out by exactly one
                                # space in the other direction, which says the
                                # mark is not there. Those are left uncompared
                                # rather than corrected by a rule that was never
                                # measured for them.
                                # ★DEFAULT: only a shape that holds ONE
                                # paragraph is compared, because that is the
                                # arrangement whose null is known. The mark
                                # correction is real -- it is one space and it
                                # doubles with the size (`gen_pptx_boundwidth.py`
                                # 14pt +2.60 against -1.27, 28pt +5.32 against
                                # -2.42) -- but a deck does not apply it as
                                # predictably as the probe does: deck 9's
                                # single-line 'THREATS' in a multi-paragraph
                                # placeholder reads -3.31pt once corrected,
                                # which says the mark was not there. So the
                                # correction lives behind `--mark-correct`
                                # until a rule for its presence is derived,
                                # and the gate keeps the population it can
                                # account for.
                                many = len(c["paras"]) > 1
                                mark = espace if (MARK_CORRECT and many and cn == 1) else 0.0
                                skip = many and (not MARK_CORRECT or cn > 1)
                                box = cw if axis == 0 else chh
                                last = min(len(box), len(ew)) - 1
                                for j, (a, b) in enumerate(zip(box, ew)):
                                    a -= mark
                                    if j != last or axis is None or skip:
                                        continue
                                    if b > 40.0:
                                        wide.append((b, a - b))
                                        if WIDTH_ROWS is not None:
                                            WIDTH_ROWS.append({
                                                "slide": si, "text": ctext[:60],
                                                "engine_pen": b, "ppt_ink": a,
                                                "delta": round(a - b, 3),
                                                "asks": cface, "measured": eface,
                                                "lines_in_para": cn,
                                            })
                                        if abs(a - b) / b > 0.02 and abs(a - b) > 1.0:
                                            wfar.append((si, round(a - b, 2),
                                                         round(100.0 * (a - b) / b, 1),
                                                         ctext[:28], cface, eface))
                        else:
                            got["diff"] += 1
                            worst.append((si, cn, en, ctext[:44]))
        finally:
            pres.Saved = True
            pres.Close()
    finally:
        app.Quit()
    rate = 100.0 * got["same"] / got["paras"] if got["paras"] else 0.0
    print(f"{str(doc):>4}: {got['paras']:5} paragraphs  {rate:6.2f}% break agreement  "
          f"({got['diff']} differ)  shapes {got['shapes']} "
          f"({got['unmatched']} unmatched, {got['turned_paras']} paragraphs "
          f"turned -- no left edge, {got['trailing_br']} end in a break)", flush=True)
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
    if wide:
        deltas = sorted(d for _, d in wide)
        med = deltas[len(deltas) // 2]
        # ★The SLOPE is the advance error and the INTERCEPT is the bearing: fit
        # `ink - pen = a * width + b` and read them apart, because a deck can
        # have a large median from side bearings alone and still measure every
        # advance correctly. Reported in per mille, since 1% on a 200pt line is
        # two points and shows in the picture.
        n = len(wide)
        sw = sum(w for w, _ in wide)
        sd = sum(d for _, d in wide)
        sww = sum(w * w for w, _ in wide)
        swd = sum(w * d for w, d in wide)
        den = n * sww - sw * sw
        slope = (n * swd - sw * sd) / den if abs(den) > 1e-9 else 0.0
        icept = (sd - slope * sw) / n
        got["width_slope_permille"] = slope * 1000.0
        got["width_intercept_pt"] = icept
        print(f"      line width: {n} lines, PowerPoint's ink minus the engine's pen "
              f"median {med:+.2f}pt; fit = {slope * 1000.0:+.2f} per mille of the "
              f"line + {icept:+.2f}pt, {len(wfar)} over 2%", flush=True)
        for si, d, pct, t, cf, ef in sorted(wfar, key=lambda r: -abs(r[2]))[:6]:
            # An EMPTY name is COM declining to answer (a line whose runs
            # disagree), not a mismatch -- flagging it as one turned d32's
            # whole listing into a false "different face".
            same = ("" if not cf or cf.lower() == ef.lower()
                    else "   <-- measured a different face")
            print(f"        s{si:<3} {d:+7.2f}pt ({pct:+.1f}%)  {t!r}  "
                  f"asks {cf!r}, measured {ef!r}{same}", flush=True)
    return got


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("docs", nargs="*", help="blind indices, or dev names like d41")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--dev", action="store_true", help="every dev deck")
    # ★Turned shapes are IN by default: their line counts are comparable and
    # they are 2.8% of the corpus's text shapes (29% of d39). This restores the
    # pre-2026-09-02 population for comparing against an older run.
    ap.add_argument("--no-rotated", action="store_true",
                    help="drop turned shapes entirely, as this did before")
    ap.add_argument("--dump-widths", default="",
                    help="write every compared line's width pair to this JSONL")
    ap.add_argument("--mark-correct", action="store_true",
                    help="also compare single-line paragraphs in multi-paragraph "
                         "shapes, taking the paragraph mark off the box")
    args = ap.parse_args()
    global ROTATED, WIDTH_ROWS, MARK_CORRECT
    ROTATED = not args.no_rotated
    MARK_CORRECT = args.mark_correct
    if args.dump_widths:
        WIDTH_ROWS = []
    docs = args.docs
    if args.all:
        manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
        docs = [i["idx"] for i in manifest]
    if args.dev:
        import re as _re
        docs = sorted({m.group(1) for f in (ROOT / "dev" / "pptx").glob("*.pptx")
                       if (m := _re.match(r"(d\d+)__", f.name))})
    if not docs:
        sys.exit("name at least one deck, or pass --all")
    tot = {"paras": 0, "same": 0, "diff": 0}
    for d in docs:
        g = audit(d)
        if g:
            for k in tot:
                tot[k] += g[k]
    if WIDTH_ROWS is not None and args.dump_widths:
        with open(args.dump_widths, "w", encoding="utf-8") as fh:
            for row in WIDTH_ROWS:
                fh.write(json.dumps(row) + "\n")
        print(f"wrote {len(WIDTH_ROWS)} width pairs to {args.dump_widths}")
    if tot["paras"]:
        print(f"\n{tot['paras']} paragraphs over {len(docs)} decks: "
              f"{100.0 * tot['same'] / tot['paras']:.2f}% break agreement, "
              f"{tot['diff']} differ")


if __name__ == "__main__":
    main()
