# -*- coding: utf-8 -*-
"""Which switch made which unseen document worse.

The A/B between two builds says the net effect on documents nobody had seen;
it does not say which of the changes did the damage. Most of this engine's
layout rules ship behind an opt-out environment variable, so each one can be
turned off on the CURRENT binary without rebuilding anything — which turns a
78-commit bisection into one render per rule per document.

    python tools/metrics/blame_flags_on_blind.py en D --docs worse

`--docs worse` takes only the documents the A/B says regressed, which is the
cheap and useful direction: it asks what broke rather than re-proving what
helped. `--docs all` is available and costs proportionally more.

A rule is reported when switching it OFF brings a document back up. That is
evidence the rule fires wrongly there, not proof — a rule can be right in
general and wrong in a case, which is the distinction this is meant to expose.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]


def main() -> int:
    if len(sys.argv) < 3 or sys.argv[1] not in ("en", "ja"):
        print(f"usage: {Path(sys.argv[0]).name} [en|ja] <letter> [--docs worse|all] [--flags FILE]")
        return 2
    lang, letter = sys.argv[1], sys.argv[2].upper()
    which = "worse"
    flags_file = Path(os.environ.get("TEMP", ".")) / "_flags_candidates.txt"
    for n, a in enumerate(sys.argv):
        if a == "--docs" and n + 1 < len(sys.argv):
            which = sys.argv[n + 1]
        if a == "--flags" and n + 1 < len(sys.argv):
            flags_file = Path(sys.argv[n + 1])

    bench = REPO / "pipeline_data" / f"{lang}_benchmark"
    module = ("_measure_ssim_blind" if lang == "en" else "_measure_ssim_jablind") + letter + "50"
    sys.path.insert(0, str(bench))
    sys.path.insert(0, str(REPO / "tools" / "metrics"))
    mod = __import__(module)
    import fitz

    ab = json.loads((mod.OUT / "_ab_binary.json").read_text(encoding="utf-8"))
    scores = {doc: v for doc, v in ab["docs"]}
    if which == "worse":
        wanted = [doc for doc, v in scores.items()
                  if v["now"]["mean"] is not None and v["then"]["mean"] is not None
                  and v["then"]["mean"] > v["now"]["mean"] + 0.0005]
    else:
        wanted = [doc for doc, v in scores.items() if v["now"]["mean"] is not None]
    by_id = {d["doc"]: d for d in mod.selections()}
    docs = [by_id[d] for d in wanted if d in by_id]

    flags = [f.strip() for f in flags_file.read_text(encoding="utf-8").split() if f.strip()]
    flags = [f for f in flags if f.endswith("_DISABLE")]
    print(f"{len(docs)} documents x {len(flags)} switches")
    if not docs or not flags:
        return 1

    truth = {}
    for d in docs:
        at = mod.WORD_PDF / f"{d['doc']}.pdf"
        if at.is_file():
            pdf = fitz.open(at)
            truth[d["doc"]] = [mod.rgb_from_pdf(pdf, i) for i in range(pdf.page_count)]
            pdf.close()

    def measure(doc: dict, folder: Path) -> float | None:
        pages = truth.get(doc["doc"])
        if not pages:
            return None
        at = folder / doc["doc"]
        n = mod.png_count(at, "p_p{}.png")
        got = []
        for i in range(min(len(pages), n)):
            got.append(mod.score(pages[i], mod.rgb_from_png(at / f"p_p{i+1}.png")))
        return round(sum(got) / len(got), 6) if got else None

    base = {d["doc"]: scores[d["doc"]]["now"]["mean"] for d in docs}
    found = []
    was_png = mod.OXI_PNG
    for n, flag in enumerate(flags, 1):
        here = mod.OUT / "flag_png"
        os.environ[flag] = "1"
        mod.OXI_PNG = here
        try:
            for d in docs:
                # Each document on its own so one slow file does not stall the rest.
                import shutil
                shutil.rmtree(here / d["doc"], ignore_errors=True)
                mod.render_oxi(d)
        finally:
            os.environ.pop(flag, None)
            mod.OXI_PNG = was_png
        moved = []
        for d in docs:
            got = measure(d, here)
            if got is None or base[d["doc"]] is None:
                continue
            if got > base[d["doc"]] + 0.01:
                moved.append((d["doc"], base[d["doc"]], got))
        if moved:
            found.append((flag, moved))
            for doc, before, after in moved:
                print(f"  [{n:3}/{len(flags)}] {flag}: {doc[:38]} {before:.4f} -> {after:.4f} "
                      f"with the rule OFF", flush=True)
        elif n % 20 == 0:
            print(f"  [{n:3}/{len(flags)}] nothing so far", flush=True)

    out = mod.OUT / "_flag_blame.json"
    out.write_text(json.dumps(
        [{"flag": f, "documents": [{"doc": d, "with_rule": b, "without_rule": a}
                                   for d, b, a in m]} for f, m in found],
        indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"\n{len(found)} switches improve at least one document when turned off")
    print(f"written: {out}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
