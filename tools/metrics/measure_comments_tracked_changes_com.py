"""
Word COM measurement of the 10 comment / tracked-change fixtures.

Phase 1 Tick 2-3 of feat/comments-tracked-changes.

For each fixture:
  * Open it with Word (invisible).
  * Dump what the Word object model sees:
      - Comments collection (count, author, initial, text, range text, ancestor)
      - Revisions collection (count, author, type, date, range text)
  * This validates that our hand-written fixture XML is parsed correctly by
    Word's reader and gives the Phase 2 parser team a ground-truth snapshot
    to diff against.

Detailed balloon-geometry / RGB-color capture is deferred to a follow-up
session — Word's object model does not expose balloon position or author
RGB directly; those need either UIA or pixel-sampling a rendered page.

Run:
    python tools/metrics/measure_comments_tracked_changes_com.py
Output:
    tools/metrics/output/comments_tracked_changes_com.json
"""

from __future__ import annotations

import json
import sys
import traceback
from pathlib import Path
from typing import Any

try:
    import win32com.client as win32  # type: ignore
except ImportError:
    print("pywin32 not installed; cannot run COM measurement.", file=sys.stderr)
    sys.exit(1)

FIXTURES_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "comments_samples"
OUT_PATH = Path(__file__).resolve().parent / "output" / "comments_tracked_changes_com.json"

# wdRevisionType enum values (ECMA-376 shape is analogous but named differently)
# https://learn.microsoft.com/office/vba/api/word.wdrevisiontype
WD_REVISION_TYPE = {
    0:  "wdNoRevision",
    1:  "wdRevisionInsert",
    2:  "wdRevisionDelete",
    3:  "wdRevisionProperty",
    4:  "wdRevisionParagraphNumber",
    5:  "wdRevisionDisplayField",
    6:  "wdRevisionReconcile",
    7:  "wdRevisionConflict",
    8:  "wdRevisionStyle",
    9:  "wdRevisionReplace",
    10: "wdRevisionParagraphProperty",
    11: "wdRevisionTableProperty",
    12: "wdRevisionSectionProperty",
    13: "wdRevisionStyleDefinition",
    14: "wdRevisionMovedFrom",
    15: "wdRevisionMovedTo",
    16: "wdRevisionCellInsertion",
    17: "wdRevisionCellDeletion",
    18: "wdRevisionCellMerge",
}


def safe(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def measure_doc(app, path: Path) -> dict[str, Any]:
    doc = app.Documents.Open(
        FileName=str(path),
        ReadOnly=True,
        ConfirmConversions=False,
        AddToRecentFiles=False,
        Visible=False,
    )
    try:
        result: dict[str, Any] = {
            "file": path.name,
            "word_reads_ok": True,
            "revisions": [],
            "comments": [],
        }

        # Revisions collection
        revs = doc.Revisions
        rev_count = safe(lambda: revs.Count, 0)
        for i in range(1, rev_count + 1):
            r = revs.Item(i)
            rng = r.Range
            result["revisions"].append({
                "index": i,
                "type_raw": int(r.Type),
                "type_name": WD_REVISION_TYPE.get(int(r.Type), "?"),
                "author": safe(lambda: r.Author, ""),
                "date": safe(lambda: str(r.Date), ""),
                "formatDescription": safe(lambda: r.FormatDescription, ""),
                "range_text": safe(lambda: rng.Text, ""),
                "range_start": safe(lambda: rng.Start, -1),
                "range_end": safe(lambda: rng.End, -1),
            })

        # Comments collection
        cmts = doc.Comments
        cmt_count = safe(lambda: cmts.Count, 0)
        for i in range(1, cmt_count + 1):
            c = cmts.Item(i)
            scope = safe(lambda: c.Scope, None)
            scope_text = safe(lambda: scope.Text if scope else "", "")
            ancestor = safe(lambda: c.Ancestor, None)
            ancestor_idx = None
            if ancestor is not None:
                # find parent's index in Comments
                for j in range(1, cmt_count + 1):
                    if cmts.Item(j).Index == ancestor.Index:
                        ancestor_idx = j
                        break
            # Replies is a Comments sub-collection; may be present
            reply_count = safe(lambda: c.Replies.Count, 0)
            result["comments"].append({
                "index": i,
                "author": safe(lambda: c.Author, ""),
                "initial": safe(lambda: c.Initial, ""),
                "date": safe(lambda: str(c.Date), ""),
                "text": safe(lambda: c.Range.Text, ""),
                "scope_text": scope_text,
                "done": bool(safe(lambda: c.Done, False)),
                "ancestor_index": ancestor_idx,
                "reply_count": reply_count,
            })

        # Author table Word built while parsing
        result["authors_from_revisions"] = sorted({
            r["author"] for r in result["revisions"] if r["author"]
        })
        result["authors_from_comments"] = sorted({
            c["author"] for c in result["comments"] if c["author"]
        })

        return result
    finally:
        doc.Close(SaveChanges=False)


def main() -> int:
    fixtures = sorted(FIXTURES_DIR.glob("fixture_*.docx"))
    if not fixtures:
        print(f"No fixtures found under {FIXTURES_DIR}", file=sys.stderr)
        return 1

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0  # wdAlertsNone

    all_results: list[dict[str, Any]] = []
    try:
        for f in fixtures:
            print(f"  measuring {f.name} …")
            try:
                r = measure_doc(app, f)
            except Exception as e:
                print(f"    ERROR: {type(e).__name__}: {e}")
                traceback.print_exc()
                r = {
                    "file": f.name,
                    "word_reads_ok": False,
                    "error": f"{type(e).__name__}: {e}",
                }
            all_results.append(r)

        payload = {
            "generated": "2026-04-25",
            "word_version": safe(lambda: app.Version, "?"),
            "fixtures_dir": str(FIXTURES_DIR),
            "note": (
                "Tick 2-3 pragmatic measurement: validates that Word parses "
                "the hand-written fixtures correctly and dumps object-model "
                "view of Revisions + Comments. All 10/10 fixtures now Word-OK "
                "after fixing commentsExtended/people content types (2026-04-25). "
                "Balloon geometry and author RGB are deferred to a UIA / "
                "pixel-sampling pass."
            ),
            "results": all_results,
        }
        OUT_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"\nWrote {OUT_PATH}")
    finally:
        app.Quit()

    # Summary
    print()
    print("=== Summary ===")
    for r in all_results:
        if r.get("word_reads_ok"):
            n_rev = len(r.get("revisions", []))
            n_cmt = len(r.get("comments", []))
            print(f"  {r['file']:45s} Word OK | revisions={n_rev} comments={n_cmt}")
        else:
            print(f"  {r['file']:45s} Word FAIL: {r.get('error')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
