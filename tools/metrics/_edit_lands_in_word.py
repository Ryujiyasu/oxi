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

That opinion is expensive. Word hangs — no dialog, no error, no end — on a
document that links to something it cannot reach, and one of the corpus's own
ORIGINALS does; its RPC endpoint also dies partway through a long run of opens,
after which every remaining file reads as unopenable. Both are handled here (a
watchdog that ends Word under the blocked call, and a restart-and-retry), but
thirty documents still take the better part of an hour.

So this is the second opinion, not the daily one. `oxi-roundtrip --sentinel`
asks the same question of every document in seconds, against our own reader:
exactly one place changed, and it holds the mark. Use this when Word's own
view of a particular document is what is in doubt.
"""

from __future__ import annotations

import argparse
import queue
import subprocess
import sys
import threading
from pathlib import Path

import win32com.client

REPO = Path(__file__).resolve().parents[2]
ORIGINALS = REPO / "tools" / "golden-test" / "documents" / "docx"


def within(seconds: float, work, *args):
    """Run `work`, or give up on it.

    A COM call cannot be interrupted from outside, and Word will sit on a
    document that links to something it cannot reach for as long as you let
    it — no dialog, no error, no end. The only lever is to end Word itself,
    which turns the blocked call into an RPC failure the worker can return
    from. So the work runs on its own thread and, if the clock runs out, Word
    is ended under it.
    """
    answer: queue.Queue = queue.Queue(maxsize=1)

    def run():
        try:
            answer.put(work(*args))
        except Exception:
            answer.put(None)

    worker = threading.Thread(target=run, daemon=True)
    worker.start()
    try:
        return answer.get(timeout=seconds), True
    except queue.Empty:
        subprocess.run(
            ["taskkill", "/F", "/IM", "WINWORD.EXE"],
            capture_output=True, check=False,
        )
        return None, False


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
    parser.add_argument("--patience", type=float, default=45.0,
                        help="seconds to give Word per document before ending it")
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
            (was, in_time) = within(args.patience, paragraphs, word, before)
            if not in_time:
                unread.append(f"{path.name} (the ORIGINAL hangs Word)")
                word = fresh_word()
                continue
            (now, in_time) = within(args.patience, paragraphs, word, path)
            if not in_time:
                wrong.append((path.name, "hangs Word, though the original does not"))
                word = fresh_word()
                continue
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
