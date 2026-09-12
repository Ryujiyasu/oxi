# -*- coding: utf-8 -*-
"""What construct starts the drift, per failing document, across every set.

A failing document's histogram says how far it slipped, not where it began.
Only the FIRST paragraph whose page differs is evidence about a cause: after
it, every later slip is the same error carried forward. So this finds that one
paragraph in each failure and reports the construct it sits in, read out of
`word/document.xml` rather than guessed from the text.

    python tools/metrics/_p1_first_slip.py en
    python tools/metrics/_p1_first_slip.py ja
    python tools/metrics/_p1_first_slip.py ja jablindD50

Counting the constructs across documents is the point: a class worth fixing is
one construct carrying many first slips, and a long tail of singletons means
the next fix has to come from somewhere else.
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import measure_pagination_oxi as MO  # noqa: E402
import pagination_diff as PD  # noqa: E402


def load_side(lang: str):
    if lang == "en":
        sys.path.insert(0, str(REPO / "pipeline_data" / "en_benchmark"))
        from _ab_env import BENCH, SETS, selected  # noqa: PLC0415
        return BENCH, SETS, selected
    sys.path.insert(0, str(REPO / "pipeline_data" / "ja_benchmark"))
    from _ja_p1 import BENCH, SETS, selected  # noqa: PLC0415
    return BENCH, SETS, selected


def constructs(path: str) -> list:
    """(paragraph ordinal, tag set) for every top-level-visible paragraph.

    Word numbers paragraphs in document order including those inside tables and
    text boxes, so the walk has to be over every `<w:p>` in the same order.
    """
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8", "replace")
    body = xml[xml.index("<w:body>"):]
    out = []
    depth_tbl = depth_txbx = 0
    pos = 0
    for m in re.finditer(r"<(/?)w:(tbl|txbxContent|p)([ />])", body):
        closing, name, _ = m.groups()
        if name == "tbl":
            depth_tbl += -1 if closing else 1
        elif name == "txbxContent":
            depth_txbx += -1 if closing else 1
        elif name == "p" and not closing:
            end = body.find("</w:p>", m.end())
            chunk = body[m.start():end if end > 0 else m.end()]
            kinds = []
            if depth_txbx > 0:
                kinds.append("textbox")
            if depth_tbl > 0:
                kinds.append("table")
            for tag, label in (("w:keepNext", "keepNext"), ("w:keepLines", "keepLines"),
                               ("w:pageBreakBefore", "pageBreakBefore"),
                               ("w:br w:type=\"page\"", "pageBreak"),
                               ("w:sectPr", "sectionBreak"), ("w:drawing", "drawing"),
                               ("w:pict", "vml"), ("instrText", "field"),
                               ("w:numPr", "numbering"), ("w:framePr", "frame"),
                               ("w:footnoteReference", "footnote")):
                if tag in chunk:
                    kinds.append(label)
            out.append((pos, kinds or ["plain"]))
            pos += 1
    return out


def main() -> int:
    lang = sys.argv[1] if len(sys.argv) > 1 else "en"
    only = sys.argv[2] if len(sys.argv) > 2 else None
    bench, sets, selected = load_side(lang)

    tally = Counter()
    rows = []
    for name, (final_name, p1) in sets.items():
        if only and name != only:
            continue
        wdir = bench / p1 / "word"
        if not wdir.is_dir():
            continue
        for doc_id, path in selected(final_name):
            truth = wdir / f"{doc_id}.json"
            if not truth.is_file():
                continue
            word = json.loads(truth.read_text(encoding="utf-8"))
            try:
                got = PD.diff_doc(doc_id, word, MO.measure_doc(path))
            except Exception:  # noqa: BLE001
                continue
            if got["pass"]:
                continue
            slips = [m for m in got.get("matches", []) if m.get("page_delta")]
            if not slips:
                rows.append((name, doc_id, -1, ["no slipped paragraph"]))
                tally["no slipped paragraph"] += 1
                continue
            first = min(slips, key=lambda m: m.get("word_i", 1 << 30))
            # Word numbers paragraphs from 1; the construct walk from 0.
            idx = first.get("word_i", 0) - 1
            kinds = dict(constructs(path)).get(idx, ["?"])
            rows.append((name, doc_id, idx, kinds))
            tally["+".join(kinds)] += 1

    print(f"{lang.upper()} first slips, {len(rows)} failing documents\n")
    print(f"{'set':12} {'doc':42} {'para':>6}  construct")
    for name, doc_id, idx, kinds in rows:
        print(f"{name:12} {doc_id:42} {idx:6}  {'+'.join(kinds)}")
    print("\nby construct:")
    for kind, n in tally.most_common():
        print(f"  {n:3}  {kind}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
