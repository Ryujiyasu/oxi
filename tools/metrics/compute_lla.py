"""Compute the Line-Layout Agreement (LLA) score between Word and Oxi.

LLA = per-doc binary pass/fail (strictest tier the project picked):
  - PASS  iff Word.n_pages == Oxi.n_pages AND
              every (page, line_idx) has Word.line == Oxi.line (exact eq).

Outputs the binary verdict, totals, and the FIRST mismatch detail so the
user can immediately localise where the gap starts. Companion to
`measure_lla_word.py` and `measure_lla_oxi.py`.

Usage:
    python compute_lla.py word.json oxi.json -o diff.json
"""
from __future__ import annotations

import argparse
import json
import os
import sys


def _lcs_len(a: list[str], b: list[str]) -> int:
    """Length of longest common subsequence — O(len_a*len_b) time/space."""
    if not a or not b:
        return 0
    prev = [0] * (len(b) + 1)
    for ai in a:
        cur = [0]
        for j, bj in enumerate(b, start=1):
            if ai == bj:
                cur.append(prev[j - 1] + 1)
            else:
                cur.append(max(prev[j], cur[-1]))
        prev = cur
    return prev[-1]


def diff_one(word: dict, oxi: dict) -> dict:
    """Per-doc diff using LCS-based sequence alignment.

    Comparing line-by-line strict index fails catastrophically when one
    side has an inserted/extra line early on (everything after that
    index shifts and counts as mismatch). LCS handles insert/delete
    naturally and reports the genuine structural agreement.

    Two scores:
      - line_text_match_rate     = LCS-aligned exact matches / max(W,O)
      - line_text_strict_rate    = index-aligned exact matches / max(W,O)
        (kept for diagnostic: high stict + lower lcs = order issue;
         low strict + higher lcs   = insertion/deletion issue)
    """
    w_pages = word.get("pages", {})
    o_pages = oxi.get("pages", {})
    w_n = word.get("n_pages", len(w_pages))
    o_n = oxi.get("n_pages", len(o_pages))
    page_count_match = (w_n == o_n)

    page_line_count_match = 0
    page_line_count_total = 0
    lcs_match = 0
    strict_match = 0
    total_lines = 0
    first_mismatch = None

    all_pages = sorted(set(w_pages) | set(o_pages), key=lambda k: int(k))
    for p in all_pages:
        wl = w_pages.get(p, [])
        ol = o_pages.get(p, [])
        page_line_count_total += 1
        if len(wl) == len(ol):
            page_line_count_match += 1
        lcs_match += _lcs_len(wl, ol)
        n = max(len(wl), len(ol))
        total_lines += n
        for i in range(n):
            w_text = wl[i] if i < len(wl) else None
            o_text = ol[i] if i < len(ol) else None
            if w_text == o_text:
                strict_match += 1
            elif first_mismatch is None:
                first_mismatch = {
                    "page": int(p),
                    "line_idx": i,
                    "word": w_text,
                    "oxi": o_text,
                }

    passed = (
        page_count_match
        and page_line_count_match == page_line_count_total
        and lcs_match == total_lines
    )

    return {
        "doc_id": word.get("doc_id") or oxi.get("doc_id"),
        "pass": passed,
        "word_pages": w_n,
        "oxi_pages": o_n,
        "page_count_match": page_count_match,
        "pages_with_matching_line_count": page_line_count_match,
        "pages_total": page_line_count_total,
        "lcs_match": lcs_match,
        "strict_match": strict_match,
        "total_lines": total_lines,
        "line_text_match_rate": lcs_match / total_lines if total_lines else 0.0,
        "line_text_strict_rate": strict_match / total_lines if total_lines else 0.0,
        "first_mismatch": first_mismatch,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("word_json", help="output of measure_lla_word.py")
    ap.add_argument("oxi_json", help="output of measure_lla_oxi.py")
    ap.add_argument("-o", "--output", default=None)
    args = ap.parse_args()

    with open(args.word_json, encoding="utf-8") as f:
        word = json.load(f)
    with open(args.oxi_json, encoding="utf-8") as f:
        oxi = json.load(f)

    result = diff_one(word, oxi)

    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(text)
        verdict = "PASS" if result["pass"] else "FAIL"
        rate = result["line_text_match_rate"]
        strict = result["line_text_strict_rate"]
        print(
            f"# {verdict}  {result['doc_id']}: "
            f"pages w={result['word_pages']} o={result['oxi_pages']}  "
            f"line-count agree {result['pages_with_matching_line_count']}/{result['pages_total']}  "
            f"LCS {result['lcs_match']}/{result['total_lines']} ({rate*100:.1f}%) "
            f"strict {result['strict_match']}/{result['total_lines']} ({strict*100:.1f}%)"
        )
        if result["first_mismatch"]:
            fm = result["first_mismatch"]
            print(f"# first mismatch p.{fm['page']} line {fm['line_idx']}:")
            print(f"#   Word: {fm['word']!r}")
            print(f"#   Oxi : {fm['oxi']!r}")
    else:
        print(text)
        sys.exit(0 if result["pass"] else 1)


if __name__ == "__main__":
    main()
