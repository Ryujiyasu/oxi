"""
Inventory scan — comment + tracked-change feature usage across the 184-doc baseline.

Phase 1 Tick 1 of feat/comments-tracked-changes mission.

Produces two JSON files in tools/metrics/output/:
  - comments_inventory.json
  - tracked_changes_inventory.json

For each .docx under oxi-main/tools/golden-test/documents/docx/:
  * Check presence of word/comments.xml + word/commentsExtended.xml + word/commentsIds.xml
  * Parse word/document.xml and count the marker element occurrences.

No Word COM needed. Pure ZIP + XML text scan.
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

DOCS_DIR = Path(r"C:/Users/ryuji/oxi-main/tools/golden-test/documents/docx")
OUT_DIR = Path(__file__).resolve().parent / "output"

# word/document.xml marker patterns. Match element open-tag only; namespaces vary.
COMMENT_PATTERNS = {
    "commentRangeStart": re.compile(r"<w:commentRangeStart\b"),
    "commentRangeEnd":   re.compile(r"<w:commentRangeEnd\b"),
    "commentReference":  re.compile(r"<w:commentReference\b"),
}

REVISION_PATTERNS = {
    "ins":        re.compile(r"<w:ins\b"),
    "del":        re.compile(r"<w:del\b"),
    "moveFrom":   re.compile(r"<w:moveFrom\b"),
    "moveTo":     re.compile(r"<w:moveTo\b"),
    "moveFromRangeStart": re.compile(r"<w:moveFromRangeStart\b"),
    "moveToRangeStart":   re.compile(r"<w:moveToRangeStart\b"),
    "rPrChange": re.compile(r"<w:rPrChange\b"),
    "pPrChange": re.compile(r"<w:pPrChange\b"),
    "tblPrChange": re.compile(r"<w:tblPrChange\b"),
    "tblGridChange": re.compile(r"<w:tblGridChange\b"),
    "trPrChange": re.compile(r"<w:trPrChange\b"),
    "tcPrChange": re.compile(r"<w:tcPrChange\b"),
    "sectPrChange": re.compile(r"<w:sectPrChange\b"),
    "numberingChange": re.compile(r"<w:numberingChange\b"),
}

# Detect unique authors by scanning w:author="..." attribute values inside comments.xml,
# and any <w:ins|del|move*... w:author="...">
AUTHOR_RE = re.compile(r'w:author="([^"]*)"')

# Comment reply thread detection: commentsExtended.xml has w:parentId="<id>" on each comment.
PARENT_ID_RE = re.compile(r'w:parentId="(\d+)"')
COMMENT_ID_RE = re.compile(r'w:id="(\d+)"')


def scan_one(docx: Path) -> dict[str, Any]:
    result: dict[str, Any] = {
        "file": docx.name,
        "has_comments_xml": False,
        "has_commentsExtended_xml": False,
        "has_commentsIds_xml": False,
        "has_peopleXml": False,
        "comment_xml_count": 0,
        "comment_reply_count": 0,
        "comment_authors": [],
        "revision_authors": [],
        "document_counts": dict.fromkeys(
            list(COMMENT_PATTERNS.keys()) + list(REVISION_PATTERNS.keys()), 0
        ),
    }
    try:
        with zipfile.ZipFile(docx, "r") as z:
            names = set(z.namelist())
            result["has_comments_xml"]         = "word/comments.xml" in names
            result["has_commentsExtended_xml"] = "word/commentsExtended.xml" in names
            result["has_commentsIds_xml"]      = "word/commentsIds.xml" in names
            result["has_peopleXml"]            = "word/people.xml" in names

            # document.xml — main marker scan
            try:
                doc = z.read("word/document.xml").decode("utf-8", errors="replace")
            except KeyError:
                doc = ""
            for label, pat in COMMENT_PATTERNS.items():
                result["document_counts"][label] = len(pat.findall(doc))
            for label, pat in REVISION_PATTERNS.items():
                result["document_counts"][label] = len(pat.findall(doc))

            # revision authors from document.xml
            if doc:
                rev_authors = set(AUTHOR_RE.findall(doc))
                result["revision_authors"] = sorted(rev_authors)

            # comments.xml — count comment bodies + collect authors
            if result["has_comments_xml"]:
                try:
                    cx = z.read("word/comments.xml").decode("utf-8", errors="replace")
                except KeyError:
                    cx = ""
                # Count top-level <w:comment ...> (excluding the wrapper).
                result["comment_xml_count"] = len(re.findall(r"<w:comment\s", cx))
                result["comment_authors"]   = sorted(set(AUTHOR_RE.findall(cx)))

            # commentsExtended.xml — count parentId markers (reply thread depth proxy)
            if result["has_commentsExtended_xml"]:
                try:
                    cex = z.read("word/commentsExtended.xml").decode(
                        "utf-8", errors="replace"
                    )
                except KeyError:
                    cex = ""
                result["comment_reply_count"] = len(PARENT_ID_RE.findall(cex))

    except zipfile.BadZipFile:
        result["error"] = "BadZipFile"
    except Exception as e:
        result["error"] = f"{type(e).__name__}: {e}"
    return result


def main() -> int:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    files = sorted(DOCS_DIR.glob("*.docx"))
    print(f"Scanning {len(files)} docx files in {DOCS_DIR}")

    per_doc: list[dict[str, Any]] = []
    for f in files:
        r = scan_one(f)
        per_doc.append(r)

    # Aggregate — comments
    comments_docs = [r for r in per_doc if r.get("has_comments_xml")]
    total_comments = sum(r["comment_xml_count"] for r in comments_docs)
    total_replies  = sum(r["comment_reply_count"] for r in comments_docs)
    all_authors: Counter[str] = Counter()
    for r in comments_docs:
        for a in r["comment_authors"]:
            all_authors[a] += 1

    comments_summary = {
        "baseline_total_docs": len(per_doc),
        "docs_with_word_comments_xml": len(comments_docs),
        "docs_with_commentsExtended": sum(
            1 for r in per_doc if r.get("has_commentsExtended_xml")
        ),
        "docs_with_commentsIds": sum(
            1 for r in per_doc if r.get("has_commentsIds_xml")
        ),
        "docs_with_peopleXml": sum(1 for r in per_doc if r.get("has_peopleXml")),
        "total_comment_bodies": total_comments,
        "total_reply_markers": total_replies,
        "authors_global": dict(all_authors.most_common()),
        "per_doc": [
            {
                "file": r["file"],
                "comment_xml_count": r["comment_xml_count"],
                "comment_reply_count": r["comment_reply_count"],
                "comment_authors": r["comment_authors"],
                "range_start_count": r["document_counts"]["commentRangeStart"],
                "range_end_count":   r["document_counts"]["commentRangeEnd"],
                "reference_count":   r["document_counts"]["commentReference"],
                "has_extended":      r.get("has_commentsExtended_xml", False),
                "has_ids":           r.get("has_commentsIds_xml", False),
                "has_people":        r.get("has_peopleXml", False),
            }
            for r in comments_docs
        ],
    }

    # Aggregate — tracked changes
    tc_docs = []
    tc_totals: Counter[str] = Counter()
    for r in per_doc:
        dc = r["document_counts"]
        rev_total = sum(dc[k] for k in REVISION_PATTERNS)
        if rev_total > 0:
            tc_docs.append(r)
            for k in REVISION_PATTERNS:
                tc_totals[k] += dc[k]

    tc_authors: Counter[str] = Counter()
    for r in tc_docs:
        for a in r["revision_authors"]:
            tc_authors[a] += 1

    tc_summary = {
        "baseline_total_docs": len(per_doc),
        "docs_with_any_revision_marker": len(tc_docs),
        "marker_totals": dict(tc_totals),
        "authors_global": dict(tc_authors.most_common()),
        "per_doc": [
            {
                "file": r["file"],
                "authors": r["revision_authors"],
                **{k: r["document_counts"][k] for k in REVISION_PATTERNS},
            }
            for r in tc_docs
        ],
    }

    (OUT_DIR / "comments_inventory.json").write_text(
        json.dumps(comments_summary, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    (OUT_DIR / "tracked_changes_inventory.json").write_text(
        json.dumps(tc_summary, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print()
    print("=== Comments inventory ===")
    print(f"  docs with word/comments.xml       : {comments_summary['docs_with_word_comments_xml']}")
    print(f"  docs with commentsExtended.xml    : {comments_summary['docs_with_commentsExtended']}")
    print(f"  docs with commentsIds.xml         : {comments_summary['docs_with_commentsIds']}")
    print(f"  docs with people.xml              : {comments_summary['docs_with_peopleXml']}")
    print(f"  total comment bodies              : {comments_summary['total_comment_bodies']}")
    print(f"  total reply parentId markers      : {comments_summary['total_reply_markers']}")
    print(f"  unique comment authors (global)   : {len(comments_summary['authors_global'])}")
    print()
    print("=== Tracked changes inventory ===")
    print(f"  docs with any revision marker     : {tc_summary['docs_with_any_revision_marker']}")
    for k, v in tc_summary["marker_totals"].items():
        if v:
            print(f"  {k:20s}: {v}")
    print(f"  unique revision authors (global)  : {len(tc_summary['authors_global'])}")
    print()
    print(f"Wrote: {OUT_DIR / 'comments_inventory.json'}")
    print(f"Wrote: {OUT_DIR / 'tracked_changes_inventory.json'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
