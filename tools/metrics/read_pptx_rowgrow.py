# -*- coding: utf-8 -*-
"""Read the row-growth probe: how tall did PowerPoint make each row?

Rows in the probe are identical and declared 1pt, so the baseline PITCH down a
column IS the grown row height — no model in between. Reports it next to the
candidates: Oxi's `2*margin + 1.2*size`, and `2*margin + (asc+desc)*size` from
the face's own OS/2 table.

Usage: python tools/metrics/read_pptx_rowgrow.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DIR = Path(r"pipeline_data\pptx_probes\rowgrow").resolve()
FONT_FILES = {
    "Arial": r"C:\Windows\Fonts\arial.ttf",
    "Segoe Script": r"C:\Windows\Fonts\segoesc.ttf",
    "Comic Sans MS": r"C:\Windows\Fonts\comic.ttf",
}


def face_metrics(name: str) -> dict[str, float] | None:
    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        return None
    path = FONT_FILES.get(name)
    if not path:
        return None
    f = TTFont(path, lazy=True)
    upm = f["head"].unitsPerEm
    os2 = f["OS/2"]
    hhea = f["hhea"]
    return {
        "win": (os2.usWinAscent + os2.usWinDescent) / upm,
        "hhea": (hhea.ascent - hhea.descent) / upm,
        "hhea_gap": (hhea.ascent - hhea.descent + hhea.lineGap) / upm,
        "typo": (os2.sTypoAscender - os2.sTypoDescender + os2.sTypoLineGap) / upm,
    }


def main() -> None:
    manifest = json.loads((DIR / "rowgrow_manifest.json").read_text(encoding="utf-8"))
    pdf = pymupdf.open(DIR / "deck.pdf")
    print(f"{'arm':34s} {'grown':>8s} {'text part':>10s} {'/size':>7s}   candidates")
    for row in manifest:
        page = pdf[row["slide"] - 1]
        ys = []
        for b in page.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for line in b["lines"]:
                for s in line["spans"]:
                    ch = s["chars"]
                    txt = "".join(c["c"] for c in ch)
                    if txt.startswith("Hxy") and abs(s["size"] - row["size"]) < 0.2:
                        ys.append(round(ch[0]["origin"][1], 2))
        ys = sorted(set(ys))
        if len(ys) < 3:
            print(f"{row['face']} {row['size']}pt: only {len(ys)} baselines found")
            continue
        # With multi-line cells the pitch alternates line-gap / row-gap; the row
        # pitch is the distance between the SAME line of consecutive rows.
        step = row["lines"]
        pitches = [ys[i + step] - ys[i] for i in range(len(ys) - step)]
        pitches = [p for p in pitches if p > 0]
        grown = sum(pitches) / len(pitches)
        text_part = grown - 2 * row["margin"]
        per_line = text_part / row["lines"]
        label = (f"{row['face']} {row['size']:.0f}pt mar{row['margin']:.1f} "
                 f"x{row['lines']}")
        m = face_metrics(row["face"]) or {}
        cands = "  ".join(
            f"{k}={v * row['size'] * row['lines'] + 2 * row['margin']:.2f}"
            for k, v in m.items()
        )
        print(f"{label:34s} {grown:8.2f} {text_part:10.2f} "
              f"{per_line / row['size']:7.4f}   oxi1.2="
              f"{1.2 * row['size'] * row['lines'] + 2 * row['margin']:.2f}  {cands}")
        # Where the first baseline sits inside its row. The row is grown to
        # exactly 2*margin + 1.2*size*lines, so a centred block starts at
        # row_top + margin and the offset below that IS the model's A.
        table_top = 36.0  # the probe places every table at 457200 EMU
        a_values = [
            (y - (table_top + i * grown) - row["margin"]) / row["size"]
            for i, y in enumerate(ys[:: row["lines"]])
        ]
        print(f"{'':34s} first baseline A = "
              + " ".join(f"{a:.4f}" for a in a_values[:4]))


if __name__ == "__main__":
    main()
