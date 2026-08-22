# -*- coding: utf-8 -*-
r"""What Excel draws for a face it has not got, and what makes it choose.

Two earlier readings were both wrong, in ways that only showed when the
question was asked twice:

* SX47 read `AR P丸ゴシック体E`'s substitute off a band that also held the
  sheet's corner cell, and concluded that the bare name draws differently
  from the same name carrying its PANOSE.
* `_xlsx_missing_face_which.py` then found bare and dressed identical — but
  both arms named the same face **in one workbook**, and Excel resolves a
  name once per document, so the bare arm was wearing the dressed arm's
  answer.

So every arm here gets a **workbook of its own**: one shape, one name, one
dressing. A separate ruler workbook holds every installed Japanese face, and
each arm is identified by the face whose ink is the same ink.

    python tools\metrics\_xlsx_missing_face_map.py
    python tools\metrics\_xlsx_missing_face_map.py --reuse
"""
import argparse
import importlib
import os
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
probe = importlib.import_module("_xlsx_missing_face_panose")
which = importlib.import_module("_xlsx_missing_face_which")

SCRATCH = Path(r"C:\tmp\xlsx_missing_face_map")

# The two the corpus asks for carry a PANOSE of their own; the rest are named
# to ask what else the answer turns on — the same vendors, other vendors, a
# Microsoft face that is simply not installed, and one that never existed.
PANOSE = {
    "AR P丸ゴシック体E": "020F0900000000000000",
    "ＤＦ特太ゴシック体": "020B0509000000000000",
    "AR 丸ゴシック体M": "020F0500000000000000",
    "リュウミン R-KL": "02020500000000000000",
}
ASKED = ["AR P丸ゴシック体E", "Nonesuch Gothic ZZ"]
# What a run can carry beside the name, one thing at a time.
DRESSINGS = [
    ("bare", ""),
    ("panose", ' panose="{panose}"'),
    ("charset jp", ' charset="-128"'),
    ("charset ansi", ' charset="0"'),
    ("pitch 50", ' pitchFamily="50"'),
    ("jp+50", ' pitchFamily="50" charset="-128"'),
    ("jp+49", ' pitchFamily="49" charset="-128"'),
    ("jp+34", ' pitchFamily="34" charset="-128"'),
    ("jp+18", ' pitchFamily="18" charset="-128"'),
    ("jp+0", ' pitchFamily="0" charset="-128"'),
]


def one_book(path, face, extra):
    probe.CASES = [("only", face, extra)]
    probe.BOOK = path
    probe.SCRATCH = SCRATCH
    probe.build()


def arms():
    held = []
    for name in ASKED:
        for dressing, extra in DRESSINGS:
            if "{panose}" in extra and name not in PANOSE:
                continue
            held.append((name, dressing, extra.format(panose=PANOSE.get(name, ""))))
    return held


def ink_of_book(path, cases, lane_and_edges=None):
    """Each shape's ink in a workbook's picture, cropped to itself."""
    picture = path.with_suffix(".excel.png")
    if not picture.exists():
        return [None] * cases
    truth = np.asarray(Image.open(picture).convert("L"))
    told = subprocess.run(
        [str(probe.RENDERER), str(path), str(path.with_suffix(".oxi.png")), "96"],
        capture_output=True, timeout=900,
        env=dict(os.environ, OXI_XLSX_DUMP_ROWS="1", OXI_XLSX_DUMP_COLUMNS="1",
                 OXI_XLSX_SHAPE_TEXT="1"))
    heights, lane = {}, 0
    for line in told.stdout.decode("utf-8", "replace").splitlines():
        parts = line.split()
        if len(parts) == 4 and parts[0] == "row":
            heights[int(parts[1])] = int(float(parts[3]))
        if len(parts) == 4 and parts[0] == "column" and lane == 0:
            lane = int(float(parts[3]))
    edges, at = {}, 0
    for index in sorted(heights):
        edges[index] = at
        at += heights[index]
    inks = []
    for index in range(cases):
        top = edges.get(index * probe.SPACING + 1, 0)
        foot = min(edges.get((index + 1) * probe.SPACING + 1, truth.shape[0]),
                   truth.shape[0])
        inks.append(which.ink_of(truth, top, foot, left=lane))
    return inks


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse", action="store_true")
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    SCRATCH.mkdir(parents=True, exist_ok=True)

    candidates = which.japanese_faces()
    ruler = SCRATCH / "ruler.xlsx"
    books = [(ruler, None)]
    for index, (name, dressing, extra) in enumerate(arms()):
        books.append((SCRATCH / f"arm{index:02d}.xlsx", (name, dressing, extra)))

    if not args.reuse:
        probe.CASES = [(face, face, "") for face in candidates]
        probe.BOOK = ruler
        probe.SCRATCH = SCRATCH
        probe.build()
        for path, held in books[1:]:
            one_book(path, held[0], held[2])
        listing = SCRATCH / "_batch.txt"
        lines = []
        for path, _ in books:
            path.with_suffix(".excel.png").unlink(missing_ok=True)
            lines.append(str(path.resolve()) + "\t"
                         + str(path.with_suffix(".excel.png").resolve()))
        listing.write_text("\n".join(lines), encoding="utf-8-sig")
        subprocess.run(["powershell", "-NoProfile", "-File", str(probe.SHOOTER),
                        "-ListFile", str(listing.resolve())],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=3600)
        listing.unlink(missing_ok=True)

    probe.CASES = [(face, face, "") for face in candidates]
    probe.BOOK = ruler
    ruled = ink_of_book(ruler, len(candidates))

    print(f"{'asked for':<22}{'dressing':<10}{'Excel draws':<40}{'differs':>8}")
    for path, held in books[1:]:
        name, dressing, _extra = held
        probe.CASES = [("only", name, "")]
        probe.BOOK = path
        mine = ink_of_book(path, 1)[0]
        if mine is None:
            print(f"{name:<22}{dressing:<10}(nothing drawn)")
            continue
        scored = []
        for face, ink in zip(candidates, ruled):
            told = which.unlike(mine, ink)
            if told is not None:
                scored.append((told[0], face))
        scored.sort()
        exact = [face for differs, face in scored if differs == 0]
        drawn = " / ".join(exact[:4]) if exact else scored[0][1]
        print(f"{name:<22}{dressing:<10}{drawn:<40}{scored[0][0]:>8}")


if __name__ == "__main__":
    main()
