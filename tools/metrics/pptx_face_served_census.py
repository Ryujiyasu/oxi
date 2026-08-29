# -*- coding: utf-8 -*-
"""Does Oxi draw each run in the family the run asked for?

`slot_face_name` registers every embedded upright part under a second, slot-
unique GDI family -- `"<typeface> #R"` / `"#B"` / `"#I"` / `"#BI"` -- so a part can be addressed by
its SLOT instead of letting GDI's weight matching choose. That is correct as far
as it goes, but it also puts names into GDI's font table that no run ever asks
for, and **GDI does not fail a name it cannot find: it picks the closest one**.

d15 is the case that showed it (2026-08-29, under `OXI_SLOTNAT_ENABLE`): a run
asking for family `"Barlow"` at weight 700 was served `"Barlow Light #B"` --
a DIFFERENT family, whose outlines are Light -- because GDI split the registered
name into `Barlow` + `Light` and matched on the first token. The title lost its
bold, and no weight rule can recover it, since the face itself is wrong.

This asks the question for every run of every deck, from the renderer's own
mouth: `OXI_DRAW_DEBUG=` (empty prefix) makes it print the family it ASKED for
and the face GDI SERVED for each run. A served face that is another family's
slot alias is the sharp signal -- an alias exists only because Oxi registered
it, so landing on one that does not belong to the requested family is Oxi's own
name pollution and nothing PowerPoint would do.

Reported per deck: run count, the distinct (asked -> served) pairs, and the
mismatches split into

    ALIAS   served another family's #R/#B alias   <- Oxi's own doing, a defect
    OTHER   served some other face entirely       <- may be a legitimate GDI
                                                     substitution for a family
                                                     nobody has

Usage:
    python tools/metrics/pptx_face_served_census.py [--decks d15,d05] [--corpus dev|blind]
"""
from __future__ import annotations

import argparse
import glob
import os
import re
import subprocess
import sys
import tempfile
from collections import Counter, defaultdict
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev" / "pptx"
BLIND = REPO / "pipeline_data" / "pptx_benchmark" / "pptx"
EXE = REPO / "tools" / "oxi-pptx-renderer" / "target" / "release" / "oxi-pptx-renderer.exe"

DRAW = re.compile(r'^DRAW ".*" family="(?P<fam>[^"]*)" size=(?P<sz>[\d.]+) '
                  r'weight=(?P<w>\d+) italic=(?P<it>\w+)')
GAVE = re.compile(r'^GDI  gave face="(?P<face>[^"]*)" tmWeight=(?P<tw>\d+)')
ALIAS = re.compile(r"^(?P<base>.*) #(?P<slot>R|B|I|BI)$")


def pairs_for(pptx: Path) -> list[tuple[str, int, str, int]]:
    """(asked family, asked weight, served face, served tmWeight) per run."""
    out: list[tuple[str, int, str, int]] = []
    with tempfile.TemporaryDirectory(prefix="served_") as tmp:
        env = dict(os.environ)
        env["OXI_DRAW_DEBUG"] = ""          # empty prefix matches every run
        proc = subprocess.run(
            [str(EXE), str(pptx), str(Path(tmp) / "slide"), "150"],
            capture_output=True, text=True, errors="replace", env=env, timeout=3600,
        )
    asked: tuple[str, int] | None = None
    for line in (proc.stderr or "").splitlines():
        m = DRAW.match(line)
        if m:
            asked = (m.group("fam"), int(m.group("w")))
            continue
        m = GAVE.match(line)
        if m and asked:
            out.append((asked[0], asked[1], m.group("face"), int(m.group("tw"))))
            asked = None
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", choices=("dev", "blind"), default="dev")
    ap.add_argument("--decks", default="")
    args = ap.parse_args()

    root = DEV if args.corpus == "dev" else BLIND
    decks = sorted(root.glob("*.pptx"))
    if args.decks:
        want = set(args.decks.split(","))
        decks = [d for d in decks if d.name.split("__")[0] in want
                 or d.name[:2] in want]

    total_alias = total_other = total_runs = 0
    for deck in decks:
        name = deck.name.split("__")[0]
        rows = pairs_for(deck)
        alias: Counter = Counter()
        other: Counter = Counter()
        for fam, w, face, tw in rows:
            m = ALIAS.match(face)
            base = m.group("base") if m else face
            if base == fam:
                continue
            if m:
                alias[(fam, w, face, tw)] += 1
            else:
                other[(fam, w, face, tw)] += 1
        total_runs += len(rows)
        total_alias += sum(alias.values())
        total_other += sum(other.values())
        flag = "  <<< ALIAS LEAK" if alias else ""
        print(f"{name}: {len(rows)} runs, {sum(alias.values())} alias / "
              f"{sum(other.values())} other{flag}", flush=True)
        for (fam, w, face, tw), n in alias.most_common():
            print(f"    ALIAS  asked {fam!r} w{w}  ->  {face!r} tmWeight={tw}   x{n}")
        for (fam, w, face, tw), n in other.most_common(6):
            print(f"    other  asked {fam!r} w{w}  ->  {face!r} tmWeight={tw}   x{n}")
    print(f"\n{len(decks)} decks, {total_runs} runs: "
          f"{total_alias} served ANOTHER family's alias, {total_other} other substitutions")


if __name__ == "__main__":
    main()
