# -*- coding: utf-8 -*-
"""Turn a `pptx_line_audit_com.py` run into a ledger that outlives its console.

The audit asks PowerPoint itself where every paragraph breaks, which costs a
render plus a COM session per deck -- an hour and a half over the blind corpus --
and then prints it. A number that expensive should not live only in a scrollback:
this parses the run's output into a table sorted by what is worth fixing next,
and writes it beside the log so the next run can be diffed against it.

Three numbers per deck, in the order they matter:

    break      paragraphs where PowerPoint's line COUNT and the engine's differ.
               Categorical, cause-attributable, the pptx analogue of pagination
    left       spread of the line's left edge against PowerPoint's own, p95.
               A constant bias cancels out of it, so this is real disagreement
    advance    line-to-line step error, p95. Deck 47's doubling showed up here

    python tools/metrics/pptx_line_audit_com.py --all > audit.log
    python tools/metrics/pptx_line_audit_rank.py audit.log
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ★The tail is `(N unmatched)` in runs before 2026-09-02 and
# `(N unmatched, M paragraphs turned -- breaks only)` after it. Anchoring on the
# closing bracket would match neither reliably and the parser would report "no
# deck lines" -- an empty ledger that reads like a clean corpus.
DECK = re.compile(
    r"^\s*(?P<doc>\S+):\s+(?P<paras>\d+) paragraphs\s+(?P<rate>[\d.]+)% break agreement\s+"
    r"\((?P<differ>\d+) differ\)\s+shapes (?P<shapes>\d+) \((?P<unmatched>\d+) unmatched"
    r"(?:,\s*(?P<turned>\d+) paragraphs turned)?"
)
LEFT = re.compile(r"p95 (?P<p95>[\d.]+)pt, (?P<over>\d+) over 3pt")
ADV = re.compile(r"line advance: (?P<steps>\d+) steps.*?p95 (?P<p95>[\d.]+)pt, (?P<over>\d+) over")
PCOUNT = re.compile(r"shapes whose PARAGRAPH COUNT disagrees: (?P<n>\d+)")


def parse(text: str) -> list[dict]:
    decks: list[dict] = []
    cur: dict | None = None
    for line in text.splitlines():
        m = DECK.match(line)
        if m:
            cur = {
                "doc": m["doc"],
                "paragraphs": int(m["paras"]),
                "agreement": float(m["rate"]),
                "differ": int(m["differ"]),
                "shapes": int(m["shapes"]),
                "unmatched": int(m["unmatched"]),
                "left_p95": None,
                "left_over3": 0,
                "adv_p95": None,
                "adv_over": 0,
                "para_count_mismatch": 0,
            }
            decks.append(cur)
            continue
        if cur is None:
            continue
        if (m := LEFT.search(line)):
            cur["left_p95"] = float(m["p95"])
            cur["left_over3"] = int(m["over"])
        elif (m := ADV.search(line)):
            cur["adv_p95"] = float(m["p95"])
            cur["adv_over"] = int(m["over"])
        elif (m := PCOUNT.search(line)):
            cur["para_count_mismatch"] = int(m["n"])
    return decks


def main() -> None:
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    log = Path(sys.argv[1])
    decks = parse(log.read_text(encoding="utf-8", errors="replace"))
    if not decks:
        sys.exit("no deck lines in that log -- is it a `pptx_line_audit_com.py` run?")

    paras = sum(d["paragraphs"] for d in decks)
    differ = sum(d["differ"] for d in decks)
    turned = sum(d["turned_paras"] for d in decks)
    print(f"{len(decks)} decks, {paras} paragraphs, {differ} break differently "
          f"({100.0 * (paras - differ) / paras:.2f}% agreement)")
    print(f"{turned} of them sit in a turned shape and were compared for their "
          f"breaks only\n")

    # ★Decks with no break disagreement are not "done": a deck can agree on
    # every line count and still put every line in the wrong place, which is
    # what `left` and `advance` are for. They are printed in their own section
    # rather than dropped.
    broken = sorted((d for d in decks if d["differ"]), key=lambda d: -d["differ"])
    if broken:
        print("BREAK disagreements (PowerPoint's line count vs the engine's):")
        for d in broken:
            print(f"  {d['doc']:>4}  {d['differ']:4} of {d['paragraphs']:5}  "
                  f"{d['agreement']:6.2f}%   left p95 {d['left_p95']}  "
                  f"adv p95 {d['adv_p95']}")
    else:
        print("BREAK: every deck agrees on every paragraph's line count.")

    placed = sorted((d for d in decks if not d["differ"] and d["left_p95"] is not None),
                    key=lambda d: -(d["left_p95"] or 0))
    print("\nPLACEMENT, among decks that agree on every break (worst left-edge first):")
    for d in placed[:12]:
        print(f"  {d['doc']:>4}  left p95 {d['left_p95']:5.2f}pt ({d['left_over3']} over 3pt)"
              f"   adv p95 {d['adv_p95']}  ({d['paragraphs']} paragraphs)")

    hole = [d for d in decks if d["unmatched"] or d["para_count_mismatch"]]
    if hole:
        print("\nCOVERAGE HOLES -- shapes the audit could not compare, which are not passes:")
        for d in hole:
            print(f"  {d['doc']:>4}  {d['unmatched']} unmatched shapes, "
                  f"{d['para_count_mismatch']} with a different paragraph count")

    out = log.with_suffix(".ledger.json")
    out.write_text(json.dumps(decks, indent=2), encoding="utf-8")
    print(f"\nwrote {out}")


if __name__ == "__main__":
    main()
