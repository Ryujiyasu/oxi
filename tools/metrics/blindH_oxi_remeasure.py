# -*- coding: utf-8 -*-
"""Re-score the Oxi column of blind-H (pixel, 150 DPI) with the CURRENT DWrite
renderer, without touching the frozen result.

The frozen measurement (pipeline_data/FROZEN_BLIND_H.md, engine 17d1d0a4) is the
version-number baseline and is never overwritten. This renders the same 50+50
documents with the current renderer into ssim_blindH50/oxi_png_<tag>/, scores
them against the same word_pdf/ with the frozen script's own functions, writes
ssim_blindH50/_remeasure_<tag>.json and prints the per-arm means next to the
frozen ones.

    python tools/metrics/blindH_oxi_remeasure.py <tag> [en|ja|both] [dwrite.exe]
"""
import importlib.util
import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
SCRIPTS = {
    "en": REPO / "pipeline_data/en_benchmark/_measure_ssim_blindH50.py",
    "ja": REPO / "pipeline_data/ja_benchmark/_measure_ssim_jablindH50.py",
}


def load(lang):
    spec = importlib.util.spec_from_file_location(f"ssim_{lang}", SCRIPTS[lang])
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def run(lang, tag, dwrite):
    mod = load(lang)
    mod.OXI_PNG = mod.OUT / f"oxi_png_{tag}"
    if dwrite:
        mod.DWRITE = Path(dwrite)
    docs = mod.selections()
    print(f"[{lang}] {len(docs)} docs, renderer {mod.DWRITE}", flush=True)
    mod.pool_run("OXI", docs, mod.render_oxi)
    rows = []
    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {pool.submit(mod.measure, d): d for d in docs}
        for i, fut in enumerate(as_completed(futures), 1):
            row = fut.result()
            if row is None:
                continue
            rows.append(row)
            print(f"[{lang} SSIM {i:2}/{len(docs)}] {row['doc']} oxi={row['oxi']['common_mean']}", flush=True)
    summary = mod.aggregate(rows)
    out = mod.OUT / f"_remeasure_{tag}.json"
    out.write_text(json.dumps({"summary": summary, "docs": rows, "renderer": str(mod.DWRITE)}, indent=1),
                   encoding="utf-8")
    frozen = json.loads(mod.RESULT.read_text(encoding="utf-8"))
    fz = {r["doc"]: r["oxi"] for r in frozen["docs"]}
    new = summary["overall"]["oxi"]
    old = frozen["summary"]["overall"]["oxi"]
    print(f"\n[{lang}] Oxi common / penalized / page-count match / n")
    print(f"   frozen : {old['common_doc_mean']} / {old['penalized_doc_mean']} / {old['page_count_match']} / {frozen['summary']['overall']['n']}")
    print(f"   {tag:7s}: {new['common_doc_mean']} / {new['penalized_doc_mean']} / {new['page_count_match']} / {summary['overall']['n']}")
    moved = []
    for r in rows:
        o = fz.get(r["doc"])
        if o and o.get("common_mean") is not None and r["oxi"]["common_mean"] is not None:
            d = r["oxi"]["common_mean"] - o["common_mean"]
            if abs(d) >= 0.005 or r["oxi"]["page_delta"] != o["page_delta"]:
                moved.append((d, r["doc"], o["common_mean"], r["oxi"]["common_mean"], o["page_delta"], r["oxi"]["page_delta"]))
    for d, doc, a, b, pa, pb in sorted(moved):
        print(f"   {d:+.4f} {doc} {a} -> {b}  pcd {pa} -> {pb}")
    print(f"   written: {out}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    tag = sys.argv[1]
    langs = ("en", "ja") if len(sys.argv) < 3 or sys.argv[2] == "both" else (sys.argv[2],)
    dwrite = sys.argv[3] if len(sys.argv) > 3 else None
    for lang in langs:
        run(lang, tag, dwrite)
