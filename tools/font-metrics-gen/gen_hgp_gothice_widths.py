# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
"""Extract the proportional advances of HGPGothicE from the installed collection.

Writes {codepoint: advance} in 2048 units/em (the scale of the existing
HGPｺﾞｼｯｸE entry in font_metrics_compact.json). Only numbers are emitted;
the font program is never copied. Requires fontTools.

    python gen_hgp_gothice_widths.py C:/Windows/Fonts/HGRGE.TTC \
        ../../crates/oxidocs-core/src/font/data/hgp_gothice_widths.json
"""

import argparse
import json
from pathlib import Path

from fontTools.ttLib import TTCollection


def extract(collection_path):
    collection = TTCollection(collection_path)
    try:
        for face in collection.fonts:
            if face["name"].getDebugName(1) != "HGPGothicE":
                continue
            scale = 2048 / face["head"].unitsPerEm
            cmap, metrics = face.getBestCmap(), face["hmtx"].metrics
            return {str(cp): round(metrics[glyph][0] * scale) for cp, glyph in sorted(cmap.items())}
    finally:
        collection.close()
    raise ValueError("HGPGothicE not found in the collection")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("collection", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    args.output.write_text(json.dumps(extract(args.collection), separators=(",", ":")), encoding="utf-8")


if __name__ == "__main__":
    main()
