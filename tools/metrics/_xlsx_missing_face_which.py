# -*- coding: utf-8 -*-
r"""Which installed face is Excel's answer for a missing one — by the pixels.

`_xlsx_missing_face_panose.py` identified `AR P丸ゴシック体E` + its PANOSE as
游ゴシック by ink width and line pitch, against five candidates. That is not an
identification: a rounded gothic can share a width and a pitch with a square
one, and the corpus has ＨＧ丸ｺﾞｼｯｸ installed. This puts every Japanese face on
the machine into one sheet beside the missing one and compares the **bitmaps**
Excel draws, so the answer is the face whose ink is the same ink.

    python tools\metrics\_xlsx_missing_face_which.py
    python tools\metrics\_xlsx_missing_face_which.py --reuse
    python tools\metrics\_xlsx_missing_face_which.py --face "ＤＦ特太ゴシック体" \
        --panose 020B0509000000000000
"""
import argparse
import importlib
import sys
from pathlib import Path

import numpy as np
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
probe = importlib.import_module("_xlsx_missing_face_panose")
census = importlib.import_module("_xlsx_panose_census")

SHIFT_JIS = -128 & 0xFF


def japanese_faces():
    """Every installed face that carries the Japanese charset."""
    have = census.installed()
    return sorted(face for face, charset in have.items()
                  if (charset & 0xFF) == SHIFT_JIS)


def ink_of(picture, top, foot, left=0):
    """The band's ink, cropped to itself, as a bitmap of trues.

    The scan starts past the first column: the sheet carries a letter in A1
    to give the used range a corner, and reading it as part of the first
    shape's ink makes that shape match nothing.
    """
    band = picture[top:foot, left:] < 128
    rows = np.flatnonzero(band.any(axis=1))
    columns = np.flatnonzero(band.any(axis=0))
    if not rows.size or not columns.size:
        return None
    return band[rows[0]:rows[-1] + 1, columns[0]:columns[-1] + 1]


def unlike(one, other):
    """How many pixels differ, once the two are laid on the same grid."""
    if one is None or other is None:
        return None
    height = max(one.shape[0], other.shape[0])
    width = max(one.shape[1], other.shape[1])
    padded = []
    for held in (one, other):
        room = np.zeros((height, width), dtype=bool)
        room[:held.shape[0], :held.shape[1]] = held
        padded.append(room)
    return int((padded[0] ^ padded[1]).sum()), int(padded[0].sum())


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    parser.add_argument("--face", default="AR P丸ゴシック体E",
                        help="the face the workbook asks for and has not got")
    parser.add_argument("--panose", default="020F0900000000000000")
    parser.add_argument("--pitch-family", default="50")
    parser.add_argument("--charset", default="-128")
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")

    dressed = (f' panose="{args.panose}" pitchFamily="{args.pitch_family}"'
               f' charset="{args.charset}"')
    candidates = japanese_faces()
    probe.CASES = [("asked, dressed", args.face, dressed),
                   ("asked, bare", args.face, "")]
    probe.CASES += [(face, face, "") for face in candidates]
    probe.SCRATCH = Path(r"C:\tmp\xlsx_missing_face_which")
    probe.BOOK = probe.SCRATCH / "which.xlsx"

    probe.build()
    picture = (probe.BOOK.with_suffix(".excel.png") if args.reuse
               else probe.shoot())
    if not picture.exists():
        print("Excel gave no picture")
        return
    truth = np.asarray(Image.open(picture).convert("L"))

    # Each shape hangs from a row of its own, so the renderer's own row
    # heights say where each band starts — dividing the picture evenly does
    # not, and quietly reads two shapes as one.
    import os
    import subprocess
    ours = probe.SCRATCH / "which.oxi.png"
    told = subprocess.run(
        [str(probe.RENDERER), str(probe.BOOK), str(ours), "96"],
        capture_output=True, timeout=600,
        env=dict(os.environ, OXI_XLSX_DUMP_ROWS="1", OXI_XLSX_DUMP_COLUMNS="1",
                 OXI_XLSX_SHAPE_TEXT="1"))
    heights = {}
    for line in told.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            heights[int(parts[1])] = int(float(parts[3]))
    edges, at = {}, 0
    for index in sorted(heights):
        edges[index] = at
        at += heights[index]
    bands = []
    for index in range(len(probe.CASES)):
        top = edges.get(index * probe.SPACING + 1, 0)
        foot = edges.get((index + 1) * probe.SPACING + 1, truth.shape[0])
        bands.append((top, min(foot, truth.shape[0])))

    lane = 0
    for line in told.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "column" and lane == 0:
            lane = int(float(parts[3]))
    reference = ink_of(truth, *bands[0], left=lane)
    if reference is None:
        print("the dressed face drew nothing")
        return
    print(f"asked for {args.face} panose {args.panose}")
    print(f"{'candidate':<28}{'differs':>9}{'of ink':>8}{'size':>12}")
    scored = []
    for (name, _face, _extra), (top, foot) in list(zip(probe.CASES, bands))[1:]:
        held = ink_of(truth, top, foot, left=lane)
        told = unlike(reference, held)
        if told is None:
            continue
        differs, ink = told
        scored.append((differs, name, held.shape, ink))
    for differs, name, shape, ink in sorted(scored)[:12]:
        print(f"{name:<28}{differs:>9}{ink:>8}{str(shape):>12}")
    print(f"\nthe dressed face's own ink: {reference.sum()} in {reference.shape}")


if __name__ == "__main__":
    main()
