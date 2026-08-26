# -*- coding: utf-8 -*-
r"""Does an edit that changes nothing leave the picture alone?

Three questions have been asked of the editor so far: does the IR come back
the same, does Office open what we wrote, and — for one file — is the value
still right. None of them asks whether the workbook still LOOKS the same, and
the IR only speaks for what it models. A column width, a border, a fill or a
style index that the writer drops would sail through all three.

The edit `oxi-roundtrip` applies asks for nothing to change, so the two
pictures must be identical — not close, identical. Anything else is something
the writer lost.

    oxi-roundtrip <corpus> --keep C:\tmp\edited_xlsx --quiet
    python tools\metrics\_edit_render_unchanged.py C:\tmp\edited_xlsx

Renders with our own engine on both sides, so Excel is not needed and the
comparison is exact.
"""

from __future__ import annotations

import argparse
import hashlib
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
RENDERER = REPO / "tools" / "oxi-xlsx-renderer" / "target" / "release" / "oxi-xlsx-renderer.exe"
ORIGINALS = REPO / "tools" / "golden-test" / "documents" / "xlsx"
SCRATCH = Path(r"C:\tmp\edit_render")


def drawn(book: Path, out: Path) -> str | None:
    """The picture's digest, or None if the renderer would not draw it."""
    try:
        done = subprocess.run(
            [str(RENDERER), str(book), str(out), "96"],
            capture_output=True, text=True, encoding="utf-8", timeout=600,
        )
    except subprocess.TimeoutExpired:
        return None
    if done.returncode != 0 or not out.exists():
        return None
    return hashlib.sha256(out.read_bytes()).hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("where", help="directory of files the editor wrote")
    parser.add_argument("--limit", type=int, default=0)
    args = parser.parse_args()
    SCRATCH.mkdir(parents=True, exist_ok=True)
    edited = sorted(p for p in Path(args.where).glob("*.xls*") if not p.name.startswith("~$"))
    if args.limit:
        edited = edited[: args.limit]

    same, moved, unread = 0, [], []
    for path in edited:
        before = ORIGINALS / path.name
        if not before.exists():
            continue
        one = drawn(before, SCRATCH / "before.png")
        two = drawn(path, SCRATCH / "after.png")
        if one is None or two is None:
            unread.append(path.name)
            continue
        if one == two:
            same += 1
        else:
            moved.append(path.name)
    print(f"  {len(edited)} workbook(s): {same} draw exactly as they did")
    for name in moved[:20]:
        print(f"    !!  {name}")
    if unread:
        print(f"  {len(unread)} would not draw: {', '.join(unread[:6])}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
