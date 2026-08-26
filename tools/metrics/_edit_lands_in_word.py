# -*- coding: utf-8 -*-
r"""Does an edit to a document arrive where it was aimed, and nowhere else?

The workbook side of this question is `_edit_lands_where_aimed.py`. A document
needs its own, because Word's idea of where the text is has nothing to do with
Excel's: `Paragraphs` runs through the body and through every table cell alike,
in the order they appear, which is exactly the reading that matters — a change
that leaked into a neighbouring cell shows up as a second paragraph moving.

    oxi-roundtrip <corpus> --quiet --sentinel OXIMARK --keep C:\tmp\aimed_docx
    python tools\metrics\_edit_lands_in_word.py C:\tmp\aimed_docx

Word is the judge rather than our own reader, because the question is whether
the change reached the file as Word understands it.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
ORIGINALS = REPO / "tools" / "golden-test" / "documents" / "docx"


def fresh_word():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    # A document that links to something outside itself will sit there trying
    # to fetch it, with no dialog to dismiss and no error to catch. These are
    # the settings that stop it reaching out.
    for name, value in (
        ("UpdateLinksAtOpen", False),
        ("ConfirmConversions", False),
        ("WarnBeforeSavingPrintingSendingMarkup", False),
        ("SaveNormalPrompt", False),
    ):
        try:
            setattr(word.Options, name, value)
        except Exception:
            pass
    return word


def paragraphs(word, path: Path) -> list[str] | None:
    """Every paragraph Word finds, body and table cells alike, in order."""
    doc = None
    try:
        doc = word.Documents.Open(
            str(path), ConfirmConversions=False, ReadOnly=True,
            AddToRecentFiles=False, Visible=False,
        )
        if doc is None:
            return None
        # One call for the lot: asking paragraph by paragraph over COM costs a
        # round trip each and a long document has thousands.
        text = doc.Content.Text
        return text.split("\r")
    except Exception:
        return None
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("where", help="directory of documents the editor wrote")
    parser.add_argument("--mark", default="OXIMARK")
    parser.add_argument("--limit", type=int, default=0)
    args = parser.parse_args()

    edited = sorted(
        p for p in Path(args.where).glob("*.docx") if not p.name.startswith("~$")
    )
    if args.limit:
        edited = edited[: args.limit]

    word = fresh_word()
    landed, wrong, unread = 0, [], []
    try:
        for path in edited:
            before = ORIGINALS / path.name
            if not before.exists():
                continue
            was = paragraphs(word, before)
            now = paragraphs(word, path)
            if was is None or now is None:
                # Word's RPC endpoint dies partway through a long run of opens
                # — a hundred documents in, it simply stops answering, and
                # every file after that reads as unopenable. Start a fresh one
                # and ask again before believing it.
                try:
                    word.Quit()
                except Exception:
                    pass
                word = fresh_word()
                was = paragraphs(word, before)
                now = paragraphs(word, path)
            if was is None or now is None:
                unread.append(path.name)
                continue
            if len(was) != len(now):
                wrong.append((path.name, f"{len(was)} paragraph(s) became {len(now)}"))
                continue
            moved = [at for at, (one, two) in enumerate(zip(was, now)) if one != two]
            if len(moved) == 1 and args.mark in now[moved[0]]:
                landed += 1
            elif not moved:
                wrong.append((path.name, "the mark never arrived"))
            elif len(moved) > 1:
                wrong.append((path.name, f"{len(moved)} paragraphs moved"))
            else:
                wrong.append((path.name, f"holds {now[moved[0]][:30]!r}, not the mark"))
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    print(f"  {len(edited)} aimed edit(s): {landed} arrived where they were aimed, alone")
    for name, why in wrong[:20]:
        print(f"    !!  {name}  {why}")
    if unread:
        print(f"  {len(unread)} could not be read: {', '.join(unread[:5])}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
