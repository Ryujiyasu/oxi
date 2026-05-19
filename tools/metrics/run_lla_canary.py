"""End-to-end LLA canary on a list of docx files.

For each doc:
1. Render Oxi layout dump via oxi-gdi-renderer
2. Build Oxi LLA JSON
3. Measure Word LLA via COM
4. Compute diff

Prints a one-line verdict per doc + aggregate stats.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
GDI_EXE = os.path.join(
    ROOT, "tools", "oxi-gdi-renderer", "target", "release",
    "oxi-gdi-renderer.exe",
)


def run_one(docx_path: str, out_dir: str) -> dict:
    base = os.path.splitext(os.path.basename(docx_path))[0]
    layout_json = os.path.join(out_dir, f"{base}__oxi_layout.json")
    oxi_json = os.path.join(out_dir, f"{base}__lla_oxi.json")
    word_json = os.path.join(out_dir, f"{base}__lla_word.json")
    diff_json = os.path.join(out_dir, f"{base}__lla_diff.json")

    # 1. Oxi layout dump
    png_prefix = os.path.join(out_dir, f"{base}__oxi")
    r = subprocess.run(
        [GDI_EXE, docx_path, png_prefix, f"--dump-layout={layout_json}",
         "--exclude=text,border,shading,box,image,clip"],
        capture_output=True, text=True,
    )
    if r.returncode != 0 or not os.path.exists(layout_json):
        return {"doc_id": base, "pass": False, "error": f"oxi-render failed: {r.stderr[:200]}"}

    # 2. Oxi LLA build
    r = subprocess.run(
        [sys.executable, os.path.join(ROOT, "tools", "metrics", "measure_lla_oxi.py"),
         layout_json, "--doc-id", base, "-o", oxi_json],
        capture_output=True, text=True,
    )
    if r.returncode != 0:
        return {"doc_id": base, "pass": False, "error": f"oxi-lla failed: {r.stderr[:200]}"}

    # 3. Word LLA (PDF-based — bypasses Word COM line-iteration quirks)
    pdf_cache = os.path.join(out_dir, f"{base}.pdf")
    r = subprocess.run(
        [sys.executable, os.path.join(ROOT, "tools", "metrics", "measure_lla_word.py"),
         docx_path, "-o", word_json, "--pdf-cache", pdf_cache],
        capture_output=True, text=True,
    )
    if r.returncode != 0:
        return {"doc_id": base, "pass": False, "error": f"word-lla failed: {r.stderr[:200]}"}

    # 4. Diff
    r = subprocess.run(
        [sys.executable, os.path.join(ROOT, "tools", "metrics", "compute_lla.py"),
         word_json, oxi_json, "-o", diff_json],
        capture_output=True, text=True,
    )
    print(r.stdout.rstrip())
    with open(diff_json, encoding="utf-8") as f:
        return json.load(f)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("docx", nargs="+", help="docx file(s) to canary")
    ap.add_argument("--out-dir", default=None,
                    help="cache directory (default: ./pipeline_data/lla_<timestamp>)")
    args = ap.parse_args()

    if args.out_dir:
        os.makedirs(args.out_dir, exist_ok=True)
        out_dir = args.out_dir
    else:
        out_dir = tempfile.mkdtemp(prefix="lla_canary_")
    print(f"# cache dir: {out_dir}")

    results = [run_one(d, out_dir) for d in args.docx]

    summary = {
        "n_total": len(results),
        "n_pass": sum(1 for r in results if r.get("pass")),
        "n_fail": sum(1 for r in results if not r.get("pass")),
        "mean_line_text_match_rate": (
            sum(r.get("line_text_match_rate", 0.0) for r in results) / max(len(results), 1)
        ),
        "results": results,
    }
    print()
    print(f"== SUMMARY ==")
    print(f"  total:                {summary['n_total']}")
    print(f"  pass:                 {summary['n_pass']}")
    print(f"  fail:                 {summary['n_fail']}")
    print(f"  mean line-match rate: {summary['mean_line_text_match_rate']:.4f}")

    summary_path = os.path.join(out_dir, "_summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"  summary: {summary_path}")


if __name__ == "__main__":
    main()
