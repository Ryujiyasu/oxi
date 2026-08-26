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

Renders with our own engines on both sides, so Office is not needed and the
comparison is exact. Workbooks, documents and decks alike.
"""

from __future__ import annotations

import argparse
import hashlib
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
def tool(name: str) -> Path:
    return REPO / "tools" / name / "target" / "release" / f"{name}.exe"


# Each format has its own engine, and each writes a page at a time from a
# prefix rather than one file — so the digest has to cover every page, or a
# document that lost its last page would pass.
ENGINES = {
    ".xlsx": tool("oxi-xlsx-renderer"),
    ".xlsm": tool("oxi-xlsx-renderer"),
    ".docx": tool("oxi-dwrite-renderer"),
    ".pptx": tool("oxi-pptx-renderer"),
}
DOCUMENTS = REPO / "tools" / "golden-test" / "documents"
SCRATCH = Path(r"C:\tmp\edit_render")


def drawn(book: Path, into: Path) -> str | None:
    """A digest of every page drawn, or None if the engine would not draw it."""
    engine = ENGINES.get(book.suffix.lower())
    if engine is None or not engine.exists():
        return None
    if into.exists():
        for stale in into.iterdir():
            stale.unlink()
    into.mkdir(parents=True, exist_ok=True)
    try:
        done = subprocess.run(
            [str(engine), str(book), str(into / "page"), "96"],
            capture_output=True, text=True, encoding="utf-8", timeout=900,
        )
    except subprocess.TimeoutExpired:
        return None
    pages = sorted(into.glob("page*"))
    if done.returncode != 0 or not pages:
        return None
    digest = hashlib.sha256()
    for page in pages:
        digest.update(page.name.encode())
        digest.update(page.read_bytes())
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("where", help="directory of files the editor wrote")
    parser.add_argument("--limit", type=int, default=0)
    args = parser.parse_args()
    SCRATCH.mkdir(parents=True, exist_ok=True)
    edited = sorted(
        p
        for p in Path(args.where).iterdir()
        if p.suffix.lower() in ENGINES and not p.name.startswith("~$")
    )
    if args.limit:
        edited = edited[: args.limit]

    same, moved, unread = 0, [], []
    for path in edited:
        kind = {".xlsx": "xlsx", ".xlsm": "xlsx", ".docx": "docx", ".pptx": "pptx"}[
            path.suffix.lower()
        ]
        before = DOCUMENTS / kind / path.name
        if not before.exists():
            continue
        one = drawn(before, SCRATCH / "before")
        two = drawn(path, SCRATCH / "after")
        if one is None or two is None:
            unread.append(path.name)
            continue
        if one == two:
            same += 1
        else:
            moved.append(path.name)
    print(f"  {len(edited)} file(s): {same} draw exactly as they did")
    for name in moved[:20]:
        print(f"    !!  {name}")
    if unread:
        print(f"  {len(unread)} would not draw: {', '.join(unread[:6])}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
