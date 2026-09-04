# -*- coding: utf-8 -*-
"""Which decks can S-SKIPBOLD actually change?

The rule only bites where the borrow it removes could have happened, which is
all three of these at once:

  1. a part is SKIPPED (its family is served from the cloud cache or the
     machine) and that part honestly claims (family F, weight>=600, slant i);
  2. NO loaded part claims (F, >=600, i) -- so `pick(F, bold)` finds nothing;
  3. SOME loaded part claims (F, <600, i) -- so branch 2 had something to
     borrow and thicken.

Decks that miss any of the three render identically under the flag, which the
A/B then confirms as byte-identical arms. Running the full corpus to discover
that costs two renders a deck; this costs a zip read.

    python tools/metrics/pptx_skipped_bold_census.py [--blind]
"""
from __future__ import annotations

import argparse
import os
import re
import struct
import sys
import zipfile
from pathlib import Path

from fontTools.ttLib import TTFont

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2] / "pipeline_data" / "pptx_benchmark"
CLOUD_ROOT = Path(os.path.expandvars(
    r"%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts"))
SLOTS = ("regular", "bold", "italic", "boldItalic")


def norm(name: str) -> str:
    return "".join(c for c in name.lower() if c.isalnum())


def eot_identity(data: bytes) -> tuple[str, int, bool] | None:
    """(family, weight, italic) from the uncompressed EOT header."""
    if len(data) < 86:
        return None
    weight = struct.unpack_from("<I", data, 28)[0]
    italic = data[27] != 0
    n = struct.unpack_from("<H", data, 82)[0]
    family = data[84:84 + n].decode("utf-16-le", "replace").strip()
    return (family, weight, italic) if family else None


def served_elsewhere() -> set[str]:
    """Families the cloud cache holds; the machine's own are added by GDI."""
    out = set()
    for path in CLOUD_ROOT.rglob("*"):
        if path.suffix.lower() not in (".ttf", ".otf", ".ttc"):
            continue
        try:
            out.add(TTFont(str(path), lazy=True, fontNumber=0)["name"].getDebugName(1))
        except Exception:
            continue
    return {norm(f) for f in out if f}


def deck_rows(path: Path, served: set[str]) -> list[str]:
    """The (family, slant) requests this deck would have answered by borrowing."""
    with zipfile.ZipFile(path) as z:
        pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
        rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8", "replace")
        rmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        skipped: list[tuple[str, int, bool]] = []
        loaded: list[tuple[str, int, bool]] = []
        for blk in re.findall(r"<p:embeddedFont>(.*?)</p:embeddedFont>", pres, re.S):
            face = re.search(r'typeface="([^"]+)"', blk)
            if not face:
                continue
            skip = norm(face.group(1)) in served
            for slot, rid in re.findall(
                    r'<p:(%s)\s+r:id="([^"]+)"' % "|".join(SLOTS), blk):
                part = "ppt/" + rmap.get(rid, "").replace("../", "")
                try:
                    ident = eot_identity(z.read(part))
                except KeyError:
                    continue
                if ident:
                    (skipped if skip else loaded).append(ident)

        rows = []
        for fam, weight, ital in skipped:
            if weight < 600:
                continue
            key = norm(fam)
            if any(norm(f) == key and w >= 600 and i == ital for f, w, i in loaded):
                continue  # an honest bold is loaded; branch 2 never ran
            borrow = [f for f, w, i in loaded
                      if norm(f) == key and w < 600 and i == ital]
            if borrow:
                rows.append("%s%s: honest bold SKIPPED, branch 2 borrowed %r"
                            % (fam, " italic" if ital else "", borrow[0]))
        return sorted(set(rows))


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--blind", action="store_true")
    args = ap.parse_args()
    src = ROOT / ("pptx" if args.blind else "dev/pptx")
    served = served_elsewhere()
    reached = []
    decks = sorted(src.glob("*.pptx"))
    for path in decks:
        try:
            rows = deck_rows(path, served)
        except Exception as exc:
            print("%-44s ERROR %s" % (path.stem[:44], exc))
            continue
        if rows:
            reached.append(path.stem)
            print(path.stem[:60])
            for row in rows:
                print("      " + row)
    print("\n%d of %d decks can be reached: %s"
          % (len(reached), len(decks),
             ", ".join(s.split("__")[0] for s in reached) or "none"))


if __name__ == "__main__":
    main()
