# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Extract advance and vertical metrics from an installed Soei collection.

Only numerical metrics are emitted; font programs are never copied.
Requires fontTools. Run on Windows with the collection installed by Office.
"""

import argparse
import json
from pathlib import Path

from fontTools.ttLib import TTCollection


def extract_metrics(collection_path):
    expected = {"HGSoeiKakugothicUB", "HGPSoeiKakugothicUB", "HGSSoeiKakugothicUB"}
    rows = []
    collection = TTCollection(collection_path)
    try:
        for face in collection.fonts:
            family = face["name"].getDebugName(1)
            if family not in expected:
                continue
            os2, hhea = face["OS/2"], face["hhea"]
            cmap, metrics = face.getBestCmap(), face["hmtx"].metrics
            rows.append(dict(
                family=family,
                units_per_em=face["head"].unitsPerEm,
                ascender=hhea.ascent,
                descender=hhea.descent,
                line_gap=hhea.lineGap,
                win_ascent=os2.usWinAscent,
                win_descent=os2.usWinDescent,
                typo_ascender=os2.sTypoAscender,
                typo_descender=os2.sTypoDescender,
                typo_line_gap=os2.sTypoLineGap,
                use_typo_metrics=bool(os2.fsSelection & 128),
                widths={str(cp): metrics[glyph][0] for cp, glyph in sorted(cmap.items())},
            ))
    finally:
        collection.close()
    if {row["family"] for row in rows} != expected or len(rows) != len(expected):
        raise ValueError("The collection must contain all three Soei Kakugothic UB faces")
    return rows


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("collection", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    data = json.dumps(extract_metrics(args.collection), separators=(",", ":"))
    args.output.write_text(data, encoding="utf-8")


if __name__ == "__main__":
    main()
