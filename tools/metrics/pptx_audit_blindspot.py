# -*- coding: utf-8 -*-
"""How much text the break audit never looks at.

`pptx_line_audit_com.py` skips a shape that carries a rotation, on BOTH sides
and for a good reason: the engine answers in the shape's own frame while COM's
`BoundLeft` is slide space, so comparing them would measure the rotation rather
than the layout. But a dimension the gate cannot see is the one that breaks
there unnoticed, so the size of that hole belongs in the ledger next to the
agreement figure.

This counts it from the files alone -- no render, no COM -- so it can run while
an audit is in flight:

    text shapes            `<p:sp>` with a non-empty `<a:t>`
    of which rotated       `<a:xfrm rot="...">` with a non-zero rot
    of which in a group    a rotated group turns its children too

★The size of the hole (2026-09-02, dev + blind):

    114 decks   14002 text shapes   390 rotated = **2.8%** unaudited
                138 of those turned by their group, not by themselves
    worst decks d39 29.4%, d40 26.8%, blind 34 16.7%, d41 11.1%

So a 100% break agreement is a claim about 97.2% of the corpus's text shapes,
and two decks hold more than a quarter of theirs outside it. Comparing a turned
shape needs the engine to report slide-space line boxes (or the audit to rotate
COM's), not a looser tolerance.

    python tools/metrics/pptx_audit_blindspot.py [dev|blind|all]
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
from pptx_dump_ab import deck_paths  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SP = re.compile(r"<p:sp>.*?</p:sp>", re.S)
GRP = re.compile(r"<p:grpSp>.*?</p:grpSp>", re.S)
ROT = re.compile(r'<a:xfrm[^>]*\brot="(-?\d+)"')
TEXT = re.compile(r"<a:t>([^<]*)</a:t>")


def deck_counts(path: Path) -> tuple[int, int, int]:
    """(text shapes, rotated, rotated because a group turns them)."""
    shapes = rotated = in_group = 0
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not re.match(r"ppt/slides/slide\d+\.xml$", name):
                continue
            xml = z.read(name).decode("utf-8", "replace")
            turned_groups = [
                g for g in GRP.findall(xml)
                if any(int(v) for v in ROT.findall(g)[:1])
            ]
            for sp in SP.findall(xml):
                if not any(t.strip() for t in TEXT.findall(sp)):
                    continue
                shapes += 1
                own = any(int(v) for v in ROT.findall(sp)[:1])
                # ★A group's rotation reaches its children, and the audit drops
                # them for the same reason -- the engine reports them in the
                # group's frame. Counting only `<p:sp>`'s own `rot` would
                # under-report the hole, which is the failure this file exists
                # to avoid.
                inside = any(sp in g for g in turned_groups)
                if own or inside:
                    rotated += 1
                    if inside and not own:
                        in_group += 1
    return shapes, rotated, in_group


def main() -> None:
    spec = sys.argv[1] if len(sys.argv) > 1 else "all"
    decks = deck_paths(spec)
    if not decks:
        sys.exit("no decks selected")
    tot = rot = grp = 0
    worst: list[tuple[float, str, int, int]] = []
    for name, src in decks:
        try:
            s, r, g = deck_counts(src)
        except Exception as e:
            print(f"{name}: unreadable ({str(e)[:40]})", flush=True)
            continue
        tot += s
        rot += r
        grp += g
        if s:
            worst.append((r / s, name, r, s))
    worst.sort(reverse=True)
    print(f"{len(decks)} decks: {tot} text shapes, {rot} rotated "
          f"({100.0 * rot / tot if tot else 0:.1f}% the audit cannot compare), "
          f"{grp} of them turned by a group\n")
    print("decks with the largest blind spot:")
    for share, name, r, s in worst[:12]:
        if not r:
            break
        print(f"  {name:>4}  {r:4} of {s:4} text shapes rotated  ({100 * share:5.1f}%)")


if __name__ == "__main__":
    main()
