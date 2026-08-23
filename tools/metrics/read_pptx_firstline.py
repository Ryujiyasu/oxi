"""Read the first-baseline probe and score it against the derived rule.

Every arm's box top is 144.0pt with all four insets 0 and `anchor=t`, so the
first baseline's offset IS `baseline - 144`.

The rule (derived 2026-08-23 from this probe, 31 arms) is about the DESCENT,
not the ascent. With `P = 1.2 * fs` the natural line box and
`D0 = P - face * fs` the face's own share below the baseline:

    n <= 1:  D = max( D0 + 0.25 * P * (n - 1),  min(D0, 0.25 * P * n) )
    n >  1:  D = max( D0,                       0.25 * P * n          )
    off = P * n - D

The baseline sits its own descent above the box's bottom; that descent is capped
at a quarter of the box, and a face already deeper than the quarter gives up a
quarter of whatever the box loses. Above single spacing the quarter is a FLOOR
instead of a cap.

★Segoe Script is the arm that makes the rule falsifiable: its face is 0.8249,
below the `0.75 * 1.2` that an ascent-shaped reading of the same numbers uses as
its floor, so it is the only installed face here that can tell the two readings
apart. It reproduces d04 slide 1's 58pt Satisfy (face 0.7877), which the
ascent reading put 6.5pt low.

Usage: python tools/metrics/read_pptx_firstline.py
"""
from __future__ import annotations

import re
import struct
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
PDF = REPO / "pipeline_data" / "pptx_probes" / "firstline" / "probe_firstline.pdf"
AREA_TOP = 144.0
WINDOWS_FONTS = Path(r"C:\Windows\Fonts")
FILES = {
    "Arial": "arial.ttf",
    "Calibri": "calibri.ttf",
    "Verdana": "verdana.ttf",
    "Georgia": "georgia.ttf",
    "Segoe Script": "segoesc.ttf",
    "Comic Sans MS": "comic.ttf",
    "Courier New": "cour.ttf",
    "Segoe Print": "segoepr.ttf",
    "MV Boli": "mvboli.ttf",
    "Lucida Sans Unicode": "l_10646.ttf",
}


def metrics(family: str) -> tuple[float, float, float] | None:
    """(asc, desc, upem) the way `runtime_baseline_offset_em` reads them."""
    path = WINDOWS_FONTS / FILES.get(family, "")
    if not path.is_file():
        return None
    blob = path.read_bytes()
    count = struct.unpack(">H", blob[4:6])[0]
    tabs = {}
    for index in range(count):
        rec = 12 + 16 * index
        tabs[blob[rec:rec + 4]] = struct.unpack(">II", blob[rec + 8:rec + 16])
    head, os2 = tabs[b"head"][0], tabs[b"OS/2"][0]
    upem = struct.unpack(">H", blob[head + 18:head + 20])[0]
    u16 = lambda k: struct.unpack(">H", blob[os2 + k:os2 + k + 2])[0]
    i16 = lambda k: struct.unpack(">h", blob[os2 + k:os2 + k + 2])[0]
    if u16(62) & 0x80:
        return float(i16(68) + i16(72)), float(-i16(70)), float(upem)
    return float(u16(74)), float(u16(76)), float(upem)


def rule(face: float, fs: float, n: float) -> float:
    """The derived first-baseline offset, in points."""
    pitch = 1.2 * fs
    natural_descent = pitch - face * fs
    quarter = 0.25 * pitch
    if n <= 1.0:
        descent = max(natural_descent + quarter * (n - 1.0), min(natural_descent, quarter * n))
    else:
        descent = max(natural_descent, quarter * n)
    return pitch * n - descent


def main() -> None:
    doc = pymupdf.open(PDF)
    worst: list[float] = []
    print(f"{'arm':22s} {'measured':>9s} {'pitch':>8s} {'face':>7s} {'rule':>9s} "
          f"{'delta':>7s}  verdict")
    for page in doc:
        spans = []
        arm = None
        for block in page.get_text("rawdict")["blocks"]:
            for line in block.get("lines", []):
                # A caption can be split across spans, and a family name has
                # SPACES in it ("Segoe Script"), so join the line and match that.
                joined = "".join("".join(c["c"] for c in sp["chars"]) for sp in line["spans"])
                hit = re.match(r"arm (.+?)\|(\d+)\|(\d+)\|(\d+)", joined)
                if hit:
                    arm = (hit.group(1), int(hit.group(2)), int(hit.group(3)), int(hit.group(4)))
                    continue
                for span in line["spans"]:
                    text = "".join(c["c"] for c in span["chars"])
                    if text.startswith("Hxg"):
                        spans.append((span["origin"][1], text))
        if arm is None or not spans:
            continue
        spans.sort()
        font, pct, size, lines = arm
        n = pct / 100.0
        off = spans[0][0] - AREA_TOP
        pitch = (spans[1][0] - spans[0][0]) if len(spans) > 1 else 1.2 * size * n
        met = metrics(font)
        face = 1.2 * met[0] / (met[0] + met[1]) if met else float("nan")
        model = rule(face, size, n)
        delta = off - model
        worst.append(abs(delta))
        print(f"{font + '|' + str(pct) + '|' + str(size) + '|' + str(lines) + 'L':24s} {off:9.3f} {pitch:8.3f} "
              f"{face:7.4f} {model:9.3f} {delta:+7.3f}  "
              f"{'ok' if abs(delta) < 0.08 else 'OFF THE RULE'}")


    if worst:
        print(f"\nworst |delta| over {len(worst)} arms: {max(worst):.3f}pt "
              f"(the per-face baseline residual, all of one sign)")


if __name__ == "__main__":
    main()
