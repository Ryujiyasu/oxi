# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Extract portable Calibri italic metrics without copying font files.

Run on Windows with fontTools installed. Pass --output to choose the JSON
artifact; the same extraction can run on a Windows CI runner.
"""
import argparse
import json
import os
from pathlib import Path

from fontTools.ttLib import TTFont


def extract(font_dir):
    rows = []
    for filename, family in [("calibrii.ttf", "Calibri Italic"),
                             ("calibriz.ttf", "Calibri Bold Italic")]:
        with TTFont(font_dir / filename) as font:
            hhea, os2 = font["hhea"], font["OS/2"]
            cmap = font.getBestCmap()
            codepoints = [cp for cp in cmap if 32 <= cp <= 0x2fff or 0xff00 <= cp <= 0xffef]
            rows.append(dict(
                family=family, units_per_em=font["head"].unitsPerEm,
                ascender=hhea.ascent, descender=hhea.descent, line_gap=hhea.lineGap,
                win_ascent=os2.usWinAscent, win_descent=os2.usWinDescent,
                typo_ascender=os2.sTypoAscender, typo_descender=os2.sTypoDescender,
                typo_line_gap=os2.sTypoLineGap,
                use_typo_metrics=bool(os2.fsSelection & (1 << 7)),
                widths={str(cp): font["hmtx"].metrics[cmap[cp]][0] for cp in codepoints},
            ))
    return rows


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--font-dir", type=Path,
                        default=Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    args.output.write_text(json.dumps(extract(args.font_dir), separators=(",", ":")) + "\n",
                           encoding="utf-8")


if __name__ == "__main__":
    main()
