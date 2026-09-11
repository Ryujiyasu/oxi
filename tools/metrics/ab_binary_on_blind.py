# -*- coding: utf-8 -*-
"""Measure two engine builds on the same untouched documents.

The question this answers: did the layout work between two commits help on
documents nobody had seen, or only on the ones it was aimed at? A gate that
reached 100% on the corpus it was tuned against says nothing about that; the
same documents, rendered by both builds and scored against the same Word PDFs,
say it exactly.

    python tools/metrics/ab_binary_on_blind.py en D <path-to-other-renderer.exe>

The other build's pages go in their own folder and the result file is not
touched — this reports, it does not publish.
"""
from __future__ import annotations

import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]


def main() -> int:
    if len(sys.argv) < 4 or sys.argv[1] not in ("en", "ja"):
        print(f"usage: {Path(sys.argv[0]).name} [en|ja] <letter> <renderer.exe>")
        return 2
    lang, letter, other = sys.argv[1], sys.argv[2].upper(), Path(sys.argv[3])
    if not other.is_file():
        print(f"no renderer at {other}")
        return 1

    bench = REPO / "pipeline_data" / f"{lang}_benchmark"
    module = ("_measure_ssim_blind" if lang == "en" else "_measure_ssim_jablind") + letter + "50"
    sys.path.insert(0, str(bench))
    sys.path.insert(0, str(REPO / "tools" / "metrics"))
    mod = __import__(module)

    import fitz

    here = mod.OUT / "oxi_png_other"
    here.mkdir(parents=True, exist_ok=True)
    mine, theirs = mod.DWRITE, other
    print(f"current : {mine}")
    print(f"other   : {theirs}")

    docs = mod.selections()
    # Render with the other build into its own folder, by pointing the module's
    # renderer and output at it for the duration.
    was_dwrite, was_png = mod.DWRITE, mod.OXI_PNG
    mod.DWRITE, mod.OXI_PNG = theirs, here
    try:
        mod.pool_run("OTHER", docs, mod.render_oxi)
    finally:
        mod.DWRITE, mod.OXI_PNG = was_dwrite, was_png

    def score_one(doc: dict):
        doc_id = doc["doc"]
        truth = mod.WORD_PDF / f"{doc_id}.pdf"
        if not truth.is_file():
            return None
        pdf = fitz.open(truth)
        n_word = pdf.page_count
        out = {}
        for label, folder in (("now", mod.OXI_PNG), ("then", here)):
            at = folder / doc_id
            n = mod.png_count(at, "p_p{}.png")
            scores = []
            for i in range(min(n_word, n)):
                scores.append(mod.score(mod.rgb_from_pdf(pdf, i),
                                        mod.rgb_from_png(at / f"p_p{i+1}.png")))
            out[label] = {
                "pages": n,
                "page_delta": n - n_word,
                "mean": round(sum(scores) / len(scores), 6) if scores else None,
            }
        pdf.close()
        return doc_id, out

    found = []
    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = [pool.submit(score_one, d) for d in docs]
        for n, fut in enumerate(as_completed(futures), 1):
            got = fut.result()
            if got:
                found.append(got)
            if n % 10 == 0:
                print(f"  scored {n}/{len(docs)}", flush=True)

    both = [(doc, v["now"]["mean"], v["then"]["mean"], v["now"]["page_delta"],
             v["then"]["page_delta"])
            for doc, v in found
            if v["now"]["mean"] is not None and v["then"]["mean"] is not None]
    now = sum(b[1] for b in both) / len(both)
    then = sum(b[2] for b in both) / len(both)
    up = [b for b in both if b[1] > b[2] + 0.0005]
    down = [b for b in both if b[2] > b[1] + 0.0005]
    pages_now = sum(1 for b in both if b[3] == 0)
    pages_then = sum(1 for b in both if b[4] == 0)

    print(f"\n{len(both)} documents rendered by both builds")
    print(f"  mean SSIM   then {then:.4f}   now {now:.4f}   ({now - then:+.4f})")
    print(f"  page count  then {pages_then}/{len(both)}   now {pages_now}/{len(both)}")
    print(f"  {len(up)} documents improved, {len(down)} regressed, "
          f"{len(both) - len(up) - len(down)} unchanged")
    for doc, a, b, *_ in sorted(down, key=lambda r: r[1] - r[2])[:5]:
        print(f"    worse  {doc[:40]:40s} {b:.4f} -> {a:.4f}")
    for doc, a, b, *_ in sorted(up, key=lambda r: r[2] - r[1])[:5]:
        print(f"    better {doc[:40]:40s} {b:.4f} -> {a:.4f}")

    (mod.OUT / "_ab_binary.json").write_text(
        json.dumps({"then": str(theirs), "now": str(mine), "docs": found},
                   indent=1, ensure_ascii=False), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
