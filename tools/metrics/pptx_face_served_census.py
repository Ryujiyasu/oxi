# -*- coding: utf-8 -*-
"""Does Oxi draw each run in the family the run asked for?

`OXI_DRAW_DEBUG=` -- an EMPTY prefix, which matches every run -- makes the
renderer print, for each run, the family it ASKED for and the face GDI SERVED.
This pairs them up over a corpus and reports where the two disagree.

*The NAME of the served face is not the test*, and reading it as one is how the
first version of this tool produced six "defects" that were all correct
(2026-08-29). A part is addressed by the typeface the DECK FILED IT UNDER, and
21% of the corpus's parts hold something else: d29 files a genuine
`(Rubik, 700)` in the bold slot of `"Rubik Medium"`, so serving
`"Rubik Medium #B"` for a bold `"Rubik"` run is `resolve_part` doing exactly its
job -- and it agrees with PowerPoint, whose own PDF subset says `Rubik/Bold/700`.

So the served face is judged by the part's OWN identity, read from the EOT
header of the deck's `.fntdata` (uncompressed: italic at 27, weight at 28, a
length-prefixed UTF-16LE family at 82). A disagreement is reported only when the
identity's FAMILY differs from the family the run asked for:

    IDENT   the part serving this run says it is a different family
    OTHER   served a face that is not one of the deck's parts at all
            (a system or cloud font, or a GDI substitution)

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


def part_identities(pptx: Path) -> dict:
    """typeface it was FILED under -> the identities its parts actually hold."""
    import struct
    import zipfile

    out: dict = defaultdict(list)
    with zipfile.ZipFile(pptx) as z:
        try:
            pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
            rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8", "replace")
        except KeyError:
            return out
        rid = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        for m in re.finditer(
                r'<p:embeddedFont><p:font typeface="([^"]+)"[^>]*/>(.*?)</p:embeddedFont>',
                pres, re.S):
            typeface = m.group(1)
            for _slot, r in re.findall(
                    r'<p:(regular|bold|italic|boldItalic) r:id="([^"]+)"', m.group(2)):
                try:
                    data = z.read("ppt/" + rid[r].replace("../", ""))
                except KeyError:
                    continue
                if len(data) < 86:
                    continue
                weight = struct.unpack_from("<I", data, 28)[0]
                italic = data[27] != 0
                n = struct.unpack_from("<H", data, 82)[0]
                fam = data[84:84 + n].decode("utf-16-le", "replace")
                out[typeface].append((fam, weight, italic))
    return out


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

    total_ident = total_other = total_runs = 0
    for deck in decks:
        name = deck.name.split("__")[0]
        rows = pairs_for(deck)
        ids = part_identities(deck)
        ident = Counter()
        other = Counter()
        for fam, w, face, tw in rows:
            m = ALIAS.match(face)
            filed = m.group("base") if m else face
            if filed == fam:
                continue
            held = ids.get(filed)
            if held is None:
                other[(fam, w, face, tw)] += 1
            elif any(h[0] == fam for h in held):
                continue
            else:
                ident[(fam, w, face, tw, tuple(sorted({h[0] for h in held})))] += 1
        total_runs += len(rows)
        total_ident += sum(ident.values())
        total_other += sum(other.values())
        flag = "  <<< IDENTITY MISMATCH" if ident else ""
        print(f"{name}: {len(rows)} runs, {sum(ident.values())} identity / "
              f"{sum(other.values())} other{flag}", flush=True)
        for (fam, w, face, tw, held), n in ident.most_common():
            print(f"    IDENT  asked {fam!r} w{w}  ->  {face!r} tmWeight={tw}"
                  f"  which holds {list(held)}   x{n}")
        for (fam, w, face, tw), n in other.most_common(6):
            print(f"    other  asked {fam!r} w{w}  ->  {face!r} tmWeight={tw}   x{n}")
    print(f"\n{len(decks)} decks, {total_runs} runs: "
          f"{total_ident} served a part that is ANOTHER family, "
          f"{total_other} served a non-embedded face")


if __name__ == "__main__":
    main()
