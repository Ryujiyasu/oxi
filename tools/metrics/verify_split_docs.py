"""Targeted verify for docs with row-split events (Step 2/3 testing).

Re-renders & SSIM-compares only the ~6 docs known to have multi-page rows,
bypassing the full 177-doc baseline for faster iteration.

Output matches pipeline.verify summary style.
"""
import subprocess
import sys
import os
import json
import glob
from pathlib import Path

SPLIT_DOCS = [
    "d77a58485f16_20240705_resources_data_outline_08",
    "e3c545fac7a7_LOD_Handbook",
    "ed025cbecffb_index-23",
    "1636d28e2c46_tokumei_08_04",
    "3a4f9fbe1a83_001620506",
    "d4d126dfe1d9_tokumei_08_01-3",
]

ROOT = Path(__file__).parent.parent.parent
BASELINE = ROOT / "pipeline_data" / "ssim_baseline.json"


def main():
    with open(BASELINE, encoding="utf-8") as f:
        baseline = json.load(f)

    docx_dir = ROOT / "tools" / "golden-test" / "documents" / "docx"
    doc_paths = [str(docx_dir / f"{n}.docx") for n in SPLIT_DOCS]

    # Clear just these doc caches so Oxi re-renders
    for doc_name in SPLIT_DOCS:
        cache = ROOT / "pipeline_data" / "oxi_png" / doc_name
        if cache.exists():
            for f in cache.glob("page_*.png"):
                f.unlink()

    # Call pipeline on just these docs
    # (pipeline.main supports --files)
    cmd = [sys.executable, "-m", "pipeline.main", *doc_paths]
    print(f"Running: {' '.join(cmd)}")
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=str(ROOT))
    print("STDOUT:", r.stdout[-3000:])
    if r.returncode != 0:
        print("STDERR:", r.stderr[-1500:])
    sys.exit(r.returncode)


if __name__ == "__main__":
    main()
