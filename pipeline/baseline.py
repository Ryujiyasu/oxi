"""Compute and save SSIM baseline for all golden test documents."""

import json
import os
import sys
from pathlib import Path
from .word_renderer import render_with_word
from .oxi_renderer import render_with_oxi
from .ssim_calculator import calculate_ssim
from .config import DATA_DIR

BASELINE_PATH = os.path.join(DATA_DIR, "ssim_baseline.json")


def compute_baseline(docx_dir: str, limit: int = 0) -> dict:
    """Compute SSIM baseline for all docx files in the directory."""
    import glob

    docx_paths = sorted(glob.glob(os.path.join(docx_dir, "*.docx")))
    if limit > 0:
        docx_paths = docx_paths[:limit]

    print(f"Computing baseline for {len(docx_paths)} files...")

    word_results = render_with_word(docx_paths)
    oxi_results = render_with_oxi(docx_paths)
    scores = calculate_ssim(word_results, oxi_results)

    # Build baseline: doc_id -> {page -> score}
    baseline = {}
    for s in scores:
        doc_id = s["doc_id"]
        page = s["page"]
        if doc_id not in baseline:
            baseline[doc_id] = {}
        baseline[doc_id][str(page)] = s["ssim_score"]

    # Save
    Path(DATA_DIR).mkdir(parents=True, exist_ok=True)
    with open(BASELINE_PATH, "w", encoding="utf-8") as f:
        json.dump(baseline, f, indent=2)

    total = sum(len(pages) for pages in baseline.values())
    avg = sum(sc for pages in baseline.values() for sc in pages.values()) / total if total else 0
    print(f"Baseline saved: {BASELINE_PATH}")
    print(f"  {len(baseline)} documents, {total} pages, avg SSIM={avg:.4f}")
    return baseline


def load_baseline() -> dict:
    """Load saved baseline."""
    if not os.path.exists(BASELINE_PATH):
        return {}
    try:
        with open(BASELINE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except UnicodeDecodeError:
        with open(BASELINE_PATH, "r", encoding="cp932") as f:
            return json.load(f)


if __name__ == "__main__":
    docx_dir = os.path.join(
        os.path.dirname(__file__), "..",
        "tools", "golden-test", "documents", "docx"
    )
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    compute_baseline(os.path.abspath(docx_dir), limit=limit)
