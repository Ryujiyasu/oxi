# -*- coding: utf-8 -*-
"""Which decks have a bold run served by a face that is not bold?

S-BOLDADV only fires there, so this says which decks an A/B can move and -- more
usefully -- which it cannot, without rendering 48 decks twice to find out.

An embedded part is EOT, whose header carries the face's own weight at offset
28. A deck is FLAGGED when some family it embeds has a `p:bold` slot whose part
is not itself bold (weight < 600), or no bold slot at all, while the deck has
runs asking for bold in that family.

    python tools/metrics/pptx_boldslot_census.py
"""
from __future__ import annotations

import json
import re
import struct
import sys
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"


def eot_weight(blob: bytes) -> int | None:
    if len(blob) < 36:
        return None
    return struct.unpack_from("<I", blob, 28)[0]


def main() -> None:
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    flagged, clean, broken = [], [], []
    for item in manifest:
        doc = f"{item['idx']:02d}"
        src = ROOT / "pptx" / item["local"]
        if not src.exists():
            continue
        try:
            z = zipfile.ZipFile(src)
        except Exception:
            broken.append(doc)
            continue
        pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
        rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8", "replace")
        rid = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        weak = []
        for m in re.finditer(r"<p:embeddedFont>(.*?)</p:embeddedFont>", pres, re.S):
            blk = m.group(1)
            tf = re.search(r'typeface="([^"]+)"', blk)
            if not tf:
                continue
            b = re.search(r'<p:bold r:id="([^"]+)"', blk)
            if not b:
                weak.append((tf.group(1), "no bold slot"))
                continue
            try:
                w = eot_weight(z.read("ppt/" + rid[b.group(1)].replace("../", "")))
            except KeyError:
                weak.append((tf.group(1), "bold slot missing its part"))
                continue
            if w is not None and w < 600:
                weak.append((tf.group(1), f"bold slot weight {w}"))
        has_bold_run = any(
            'b="1"' in z.read(n).decode("utf-8", "replace")
            for n in z.namelist()
            if re.fullmatch(r"ppt/(slides|slideLayouts|slideMasters)/[^/]+\.xml", n)
        )
        (flagged if (weak and has_bold_run) else clean).append((doc, weak))
        z.close()
    print(f"{len(flagged)} decks can be moved, {len(clean)} cannot"
          + (f", {len(broken)} unreadable ({','.join(broken)})" if broken else ""))
    for doc, weak in flagged:
        print(f"  {doc}: " + "; ".join(f"{f} -- {why}" for f, why in weak[:4]))
    print("\nnot movable: " + ",".join(d for d, _ in clean))


if __name__ == "__main__":
    main()
