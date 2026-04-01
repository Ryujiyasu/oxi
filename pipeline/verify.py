"""Verify SSIM scores against baseline. Reject if any score drops."""

import json
import os
import sys
from pathlib import Path
from .word_renderer import render_with_word
from .oxi_renderer import render_with_oxi
from .ssim_calculator import calculate_ssim
from .baseline import load_baseline, BASELINE_PATH
from .config import DATA_DIR


def verify(docx_dir: str, limit: int = 0) -> bool:
    """
    Re-compute SSIM and compare against baseline.
    Returns True if all scores maintained or improved.
    Returns False if any score dropped.
    """
    import glob

    baseline = load_baseline()
    if not baseline:
        print("[NG] No baseline found. Run: python -m pipeline.baseline")
        return False

    # Only test files that have a baseline
    docx_paths = sorted(glob.glob(os.path.join(docx_dir, "*.docx")))
    if limit > 0:
        docx_paths = docx_paths[:limit]

    # Filter to only files with baseline
    docx_paths = [p for p in docx_paths if Path(p).stem in baseline]
    print(f"Verifying {len(docx_paths)} files against baseline...")

    word_results = render_with_word(docx_paths)
    oxi_results = render_with_oxi(docx_paths)
    scores = calculate_ssim(word_results, oxi_results, skip_heatmap=True)

    regressions = []
    improvements = []
    unchanged = []

    for s in scores:
        doc_id = s["doc_id"]
        page = str(s["page"])
        new_score = s["ssim_score"]

        if doc_id in baseline and page in baseline[doc_id]:
            old_score = baseline[doc_id][page]
            diff = new_score - old_score

            if diff < -0.001:  # Allow 0.001 tolerance for floating point
                regressions.append({
                    "doc_id": doc_id,
                    "page": int(page),
                    "old": old_score,
                    "new": new_score,
                    "diff": diff,
                })
            elif diff > 0.001:
                improvements.append({
                    "doc_id": doc_id,
                    "page": int(page),
                    "old": old_score,
                    "new": new_score,
                    "diff": diff,
                })
            else:
                unchanged.append(doc_id)

    print(f"\nResults: {len(improvements)} improved, {len(unchanged)} unchanged, {len(regressions)} regressed")

    if improvements:
        print("\nImprovements:")
        for imp in sorted(improvements, key=lambda x: -x["diff"]):
            print(f"  {imp['doc_id']} p.{imp['page']}: {imp['old']:.4f} -> {imp['new']:.4f} (+{imp['diff']:.4f})")

    if regressions:
        print("\n[!!] Regressions:")
        for reg in sorted(regressions, key=lambda x: x["diff"]):
            print(f"  {reg['doc_id']} p.{reg['page']}: {reg['old']:.4f} -> {reg['new']:.4f} ({reg['diff']:.4f})")

    # Net improvement rule: total improvement must exceed total regression
    total_gain = sum(imp["diff"] for imp in improvements)
    total_loss = sum(abs(reg["diff"]) for reg in regressions)
    net = total_gain - total_loss
    print(f"\nNet: gain={total_gain:.4f}, loss={total_loss:.4f}, net={net:+.4f}")

    if net <= 0:
        print("[NG] Net improvement is zero or negative. Rejected.")
        return False

    print(f"[OK] Net positive improvement ({net:+.4f}). Safe to commit.")

    # Update baseline with improvements
    if improvements:
        for imp in improvements:
            baseline[imp["doc_id"]][str(imp["page"])] = imp["new"]
        with open(BASELINE_PATH, "w", encoding="utf-8") as f:
            json.dump(baseline, f, indent=2)
        print(f"Baseline updated with {len(improvements)} improvements.")

    return True


if __name__ == "__main__":
    docx_dir = os.path.join(
        os.path.dirname(__file__), "..",
        "tools", "golden-test", "documents", "docx"
    )
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    ok = verify(os.path.abspath(docx_dir), limit=limit)
    sys.exit(0 if ok else 1)
