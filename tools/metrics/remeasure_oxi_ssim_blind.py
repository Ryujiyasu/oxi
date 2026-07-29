# -*- coding: utf-8 -*-
"""Re-measure ONLY the Oxi column of a frozen blind SSIM set, preserving every
other engine's per-doc results, then re-aggregate the 3/4/5/6-way summaries.

Why a dedicated script: the original `_measure_ssim_*.py` main() rewrites
`_result.json` with `{summary, docs}` only — running it after the multi-engine
columns were merged in DROPS oo/silurus/betteroffice/eigenpal (this happened on
2026-07-21 and had to be repaired from an archived file). This script never
touches the other engines' numbers.

  python remeasure_oxi_ssim_blind.py ja      # ja_benchmark/ssim_blind50
  python remeasure_oxi_ssim_blind.py en      # en_benchmark/ssim_blindB50

Old Oxi PNGs are archived to `oxi_png_<YYYYMMDD_HHMM>` before re-rendering, and
the result records the engine commit + measurement timestamp (the 2026-07-21
lesson: an undated result file cannot be compared across engine versions).
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]

SETS = {
    "ja": ("pipeline_data/ja_benchmark", "_measure_ssim_jablind50", "ssim_blind50"),
    "en": ("pipeline_data/en_benchmark", "_measure_ssim_blindB50", "ssim_blindB50"),
}

ENGINES = ["oxi", "libre", "oo", "silurus", "betteroffice", "eigenpal"]
WAYS = {"summary": 2, "summary_3way": 3, "summary_4way": 4,
        "summary_5way": 5, "summary_6way": 6}


def stats(items: list[dict], engines: list[str]) -> dict:
    result = {"n": len(items)}
    for eng in engines:
        common = [r[eng]["common_mean"] for r in items
                  if eng in r and r[eng] and r[eng].get("common_mean") is not None]
        penal = [r[eng]["penalized_mean"] for r in items
                 if eng in r and r[eng] and r[eng].get("penalized_mean") is not None]
        result[eng] = {
            "common_doc_mean": round(sum(common) / len(common), 6) if common else None,
            "penalized_doc_mean": round(sum(penal) / len(penal), 6) if penal else None,
            "page_count_match": sum(1 for r in items
                                    if eng in r and r[eng] and r[eng].get("page_delta") == 0),
        }
    paired = [r for r in items
              if all(eng in r and r[eng] and r[eng].get("common_mean") is not None
                     for eng in engines)]
    result[f"paired_n_{len(engines)}way" if len(engines) > 2 else "paired_n"] = len(paired)
    for eng in engines:
        if eng == "oxi":
            continue
        result[f"oxi_beats_{eng}"] = sum(
            r["oxi"]["common_mean"] > r[eng]["common_mean"] + 0.0005 for r in paired)
    if engines == ["oxi", "libre"]:
        result["oxi_common_wins"] = result.pop("oxi_beats_libre")
        result["libre_common_wins"] = sum(
            r["libre"]["common_mean"] > r["oxi"]["common_mean"] + 0.0005 for r in paired)
        result["common_ties"] = len(paired) - result["oxi_common_wins"] - result["libre_common_wins"]
    return result


def main() -> None:
    which = (sys.argv[1] if len(sys.argv) > 1 else "ja").lower()
    if which not in SETS:
        raise SystemExit(f"usage: {sys.argv[0]} [{'|'.join(SETS)}]")
    bench_rel, mod_name, out_name = SETS[which]
    bench = REPO / bench_rel
    sys.path.insert(0, str(bench))
    sys.path.insert(0, str(REPO / "tools" / "metrics"))
    mod = __import__(mod_name)

    out = bench / out_name
    result_path = out / "_result.json"
    data = json.loads(result_path.read_text(encoding="utf-8"))
    rows = data["docs"]
    by_doc = {r["doc"]: r for r in rows}
    print(f"{which}: {len(rows)} docs in {result_path.name}")

    commit = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=REPO,
                            capture_output=True, text=True).stdout.strip()
    stamp = time.strftime("%Y%m%d_%H%M")

    # 1. archive the old Oxi renders, then re-render with the current binary
    old = out / "oxi_png"
    if old.is_dir():
        archive = out / f"oxi_png_{stamp}"
        print(f"  archiving {old.name} -> {archive.name}")
        old.rename(archive)
    (out / "oxi_png").mkdir(parents=True, exist_ok=True)

    docs = [d for d in mod.selections() if d["doc"] in by_doc]
    print(f"  re-rendering Oxi for {len(docs)} docs (DWrite {mod.DWRITE.name})")
    mod.pool_run("OXI", docs, mod.render_oxi)

    # 2. re-score ONLY the Oxi column
    def rescore(doc: dict) -> tuple[str, dict] | None:
        import fitz
        doc_id = doc["doc"]
        wp = mod.WORD_PDF / f"{doc_id}.pdf"
        if not wp.is_file():
            return None
        word_pdf = fitz.open(wp)
        oxi_dir = mod.OXI_PNG / doc_id
        n_word = word_pdf.page_count
        n_oxi = mod.png_count(oxi_dir, "p_p{}.png")
        scores = []
        for i in range(n_word):
            if i >= n_oxi:
                break
            ref = mod.rgb_from_pdf(word_pdf, i)
            scores.append(mod.score(ref, mod.rgb_from_png(oxi_dir / f"p_p{i+1}.png")))
        word_pdf.close()
        denom = max(n_word, n_oxi)
        return doc_id, {
            "pages": n_oxi,
            "page_delta": n_oxi - n_word,
            "common_pages": len(scores),
            "common_mean": round(sum(scores) / len(scores), 6) if scores else None,
            "penalized_mean": round(sum(scores) / denom, 6) if denom else None,
            "page_min": round(min(scores), 6) if scores else None,
        }

    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {pool.submit(rescore, d): d for d in docs}
        for i, fut in enumerate(as_completed(futures), 1):
            res = fut.result()
            if res is None:
                continue
            doc_id, col = res
            prev = by_doc[doc_id].get("oxi", {}).get("common_mean")
            by_doc[doc_id]["oxi"] = col
            print(f"  [{i:2}/{len(docs)}] {doc_id}: oxi {prev} -> {col['common_mean']}",
                  flush=True)

    # 3. re-aggregate every N-way summary that was present, preserving structure
    groups = defaultdict(list)
    for r in rows:
        groups[r["type"]].append(r)
    for key, n in WAYS.items():
        if key not in data:
            continue
        engs = ENGINES[:n] if n > 2 else ["oxi", "libre"]
        method = data[key].get("method") if isinstance(data[key], dict) else None
        newsum = {"overall": stats(rows, engs),
                  "by_type": {k: stats(v, engs) for k, v in sorted(groups.items())}}
        if method:
            newsum = {"method": method, **newsum}
        data[key] = newsum

    m = data.get("summary", {}).get("method")
    if isinstance(m, dict):
        m["engine_commit"] = commit
        m["measured_at"] = stamp
        m["oxi_column"] = "re-measured; all other engine columns preserved verbatim"

    result_path.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    top = data.get("summary_6way", data.get("summary"))["overall"]
    print(json.dumps({k: v for k, v in top.items() if k != "by_type"}, indent=1))
    print(f"written: {result_path}  (engine {commit}, {stamp})")


if __name__ == "__main__":
    main()
