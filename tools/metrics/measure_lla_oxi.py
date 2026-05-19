"""Build Oxi-side per-page, per-line text from a `--dump-layout` JSON.

LLA (Line-Layout Agreement) metric — Oxi side.

Pipeline:
1. Caller runs `oxi-gdi-renderer --dump-layout=<out.json>` first.
2. This script reads that JSON, groups Text elements by Y position within
   each page, sorts by X, and concatenates into per-line strings.

Output schema (UTF-8 JSON):
{
  "doc_id":     "<basename>",
  "n_pages":    N,
  "pages": {
    "1": ["line_text_1", "line_text_2", ...],
    "2": [...],
    ...
  }
}

NOTE: This prototype handles MAIN-STORY (body) only. Headers/footers/
comments/footnotes are folded into the page they appear on (they share
the body's coordinate system in the dump). Detecting which Text elements
come from header vs body is a follow-up enhancement.
"""
from __future__ import annotations

import argparse
import json
import os
from collections import defaultdict

Y_BUCKET_TOL_PT = 3.0   # group Text elements with |dy| <= this into the same line


def _normalise(text: str) -> str:
    # Strip ALL whitespace, both ASCII and U+3000 (ideographic). The Word
    # PDF export silently converts docx U+3000 padding to runs of ASCII
    # spaces, while Oxi's IR preserves U+3000 verbatim. Either side may
    # also pad between adjacent cell contents with literal spaces. To
    # measure structural line agreement we treat all whitespace as
    # equivalent and strip it entirely; semantic content (kanji, kana,
    # ASCII letters, punctuation) drives the equality check.
    out = []
    for ch in text:
        if ch in (" ", "\t", "　"):
            continue
        out.append(ch)
    return "".join(out)


def lines_from_layout(layout: dict) -> dict[str, list[str]]:
    """Group Text elements by Y position into visual lines.

    Cell boundaries are intentionally NOT used as a grouping key: Word's
    PDF export joins cells sharing the same Y into one visual line (e.g.
    label cell + value cell on the same row → one line of text). To match
    that, we cluster purely by Y.

    Some Text elements have embedded `\\n` characters (soft line breaks
    that the IR didn't split). Expand them into separate logical pieces
    so they don't pollute the merged-line text.
    """
    pages_out: dict[str, list[str]] = {}
    for page_idx, page in enumerate(layout.get("pages", []), start=1):
        text_elems = [e for e in page.get("elements", []) if e.get("type") == "text"]
        text_elems.sort(key=lambda e: (e["y"], e["x"]))
        buckets: list[tuple[float, list[tuple[float, str]]]] = []
        for e in text_elems:
            y, x, t = e["y"], e["x"], e.get("text", "")
            # Strip embedded \n — these are explicit-soft-break markers that
            # the IR didn't split into separate Text elements. They make
            # the bucketed line text look like 'foo\nbar' which never
            # matches Word's per-visual-line output.
            t = t.replace("\n", "").replace("\r", "")
            if not t:
                continue
            if buckets and abs(y - buckets[-1][0]) <= Y_BUCKET_TOL_PT:
                buckets[-1][1].append((x, t))
            else:
                buckets.append((y, [(x, t)]))

        lines: list[str] = []
        for _y, chars in buckets:
            chars.sort(key=lambda c: c[0])
            lines.append(_normalise("".join(c[1] for c in chars)))
        # Drop pure-empty lines — Word's PDF export sometimes emits them
        # (empty paragraphs in cells) and sometimes doesn't (whitespace-
        # only cells). Both sides should agree on "ignore empty lines".
        lines = [ln for ln in lines if ln != ""]
        pages_out[str(page_idx)] = lines
    return pages_out


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("layout_json", help="path to oxi-gdi-renderer --dump-layout output")
    ap.add_argument("--doc-id", default=None, help="explicit doc_id label")
    ap.add_argument("-o", "--output", default=None, help="output JSON path (default: stdout)")
    args = ap.parse_args()

    with open(args.layout_json, encoding="utf-8") as f:
        layout = json.load(f)

    doc_id = args.doc_id or os.path.splitext(os.path.basename(args.layout_json))[0]
    pages = lines_from_layout(layout)

    result = {
        "doc_id": doc_id,
        "n_pages": len(pages),
        "pages": pages,
    }

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"# wrote {args.output}  ({len(pages)} pages, {sum(len(v) for v in pages.values())} lines)")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
