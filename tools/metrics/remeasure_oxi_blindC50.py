# -*- coding: utf-8 -*-
"""Re-measure ONLY the Oxi column of the frozen blind-C50 sets.

The other engines in these files were measured once, against the same Word
PDFs, and have not changed since; re-running them would cost hours and could
only reproduce what is already there. What changes is Oxi. So this re-renders
and re-scores Oxi alone and leaves every other engine's per-document numbers
exactly as they were.

    python tools/metrics/remeasure_oxi_blindC50.py ja
    python tools/metrics/remeasure_oxi_blindC50.py en

The sibling `remeasure_oxi_ssim_blind.py` does the same for the older blind50 /
blindB50 sets, whose files have a different summary shape. This one is for the
C50 sets the README quotes.

The old Oxi PNGs are archived rather than deleted, and the result records which
engine commit produced the new ones — an undated result cannot be compared
across engine versions, which is the whole reason for measuring again.
"""
from __future__ import annotations

import json
import subprocess
import sys
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]

SETS = {
    "ja": ("pipeline_data/ja_benchmark", "_measure_ssim_jablindC50"),
    "en": ("pipeline_data/en_benchmark", "_measure_ssim_blindC50"),
}


def wins(rows: list[dict], overall: dict) -> dict:
    """Add back the head-to-head counts the stored summary carries.

    The per-engine MEANS are not in the summary — they are computed from `docs`
    when they are wanted — but the win counts are, so they have to be rebuilt
    or the file would come back poorer than it went in.
    """
    others = [key for key in rows[0]
              if key not in ("doc", "batch", "type", "word_pages", "oxi", "libre")]
    for engine in ["libre", *others]:
        paired = [r for r in rows
                  if r.get("oxi", {}).get("common_mean") is not None
                  and r.get(engine) and r[engine].get("common_mean") is not None]
        if engine != "libre":
            overall[f"oxi_beats_{engine}"] = sum(
                r["oxi"]["common_mean"] > r[engine]["common_mean"] + 0.0005 for r in paired)
        overall[f"paired_n_{engine}"] = len(paired)
    return overall


def main() -> int:
    which = (sys.argv[1] if len(sys.argv) > 1 else "").lower()
    if which not in SETS:
        print(f"usage: {Path(sys.argv[0]).name} [{'|'.join(SETS)}]")
        return 2
    bench_rel, mod_name = SETS[which]
    bench = REPO / bench_rel
    sys.path.insert(0, str(bench))
    sys.path.insert(0, str(REPO / "tools" / "metrics"))
    mod = __import__(mod_name)

    result_path = mod.OUT / "_result.json"
    data = json.loads(result_path.read_text(encoding="utf-8"))
    rows = data["docs"]
    by_doc = {r["doc"]: r for r in rows}
    print(f"{which}: {len(rows)} documents in {result_path.parent.name}")

    if not mod.DWRITE.is_file():
        print(f"the renderer is not built: {mod.DWRITE}")
        return 1
    built = time.strftime("%Y-%m-%d %H:%M", time.localtime(mod.DWRITE.stat().st_mtime))
    commit = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=REPO,
                            capture_output=True, text=True).stdout.strip()
    stamp = time.strftime("%Y%m%d_%H%M")
    print(f"  renderer built {built}; engine at {commit}")

    old = mod.OXI_PNG
    if old.is_dir():
        archive = old.parent / f"oxi_png_{stamp}"
        print(f"  archiving {old.name} -> {archive.name}")
        old.rename(archive)
    mod.OXI_PNG.mkdir(parents=True, exist_ok=True)

    docs = [d for d in mod.selections() if d["doc"] in by_doc]
    print(f"  re-rendering {len(docs)} documents", flush=True)
    mod.pool_run("OXI", docs, mod.render_oxi)

    def rescore(doc: dict):
        import fitz
        doc_id = doc["doc"]
        word = mod.WORD_PDF / f"{doc_id}.pdf"
        if not word.is_file():
            return None
        pdf = fitz.open(word)
        here = mod.OXI_PNG / doc_id
        n_word = pdf.page_count
        n_oxi = mod.png_count(here, "p_p{}.png")
        scores = []
        for i in range(min(n_word, n_oxi)):
            scores.append(mod.score(mod.rgb_from_pdf(pdf, i),
                                    mod.rgb_from_png(here / f"p_p{i+1}.png")))
        pdf.close()
        denom = max(n_word, n_oxi)
        return doc_id, {
            "pages": n_oxi,
            "page_delta": n_oxi - n_word,
            "common_pages": len(scores),
            "common_mean": round(sum(scores) / len(scores), 6) if scores else None,
            "penalized_mean": round(sum(scores) / denom, 6) if denom else None,
            "page_min": round(min(scores), 6) if scores else None,
        }

    moved = []
    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {pool.submit(rescore, d): d for d in docs}
        for n, fut in enumerate(as_completed(futures), 1):
            got = fut.result()
            if got is None:
                continue
            doc_id, column = got
            before = by_doc[doc_id].get("oxi", {}).get("common_mean")
            by_doc[doc_id]["oxi"] = column
            if before is not None and column["common_mean"] is not None:
                moved.append((doc_id, before, column["common_mean"]))
            print(f"  [{n:2}/{len(docs)}] {doc_id}: {before} -> {column['common_mean']}",
                  flush=True)

    groups = defaultdict(list)
    for row in rows:
        groups[row["type"]].append(row)
    fresh = mod.aggregate(rows)
    data["summary"] = {
        "method": {**fresh["method"], "engine_commit": commit, "measured_at": stamp,
                   "oxi_column": "re-measured; every other engine preserved verbatim"},
        "overall": wins(rows, fresh["overall"]),
        "by_type": fresh["by_type"],
    }
    result_path.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")

    up = sum(1 for _, a, b in moved if b > a + 0.0005)
    down = sum(1 for _, a, b in moved if a > b + 0.0005)
    overall = data["summary"]["overall"]
    print(f"\n  Oxi common mean {overall['oxi']['common_doc_mean']:.4f}"
          f"   penalized {overall['oxi']['penalized_doc_mean']:.4f}"
          f"   pages match {overall['oxi']['page_count_match']}/{overall['n']}")
    print(f"  {up} documents improved, {down} regressed, "
          f"{len(moved) - up - down} unchanged")
    for doc_id, a, b in sorted(moved, key=lambda m: m[2] - m[1])[:5]:
        if a - b > 0.0005:
            print(f"    down  {doc_id}  {a:.4f} -> {b:.4f}")
    print(f"  written: {result_path}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
