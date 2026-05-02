"""Bottom-15 canary — fast SSIM check on the 15 worst docs only.

Stages bottom-15 baseline docs into a temp dir and runs pipeline.verify.
Reports per-page diffs without touching the baseline. Useful for canarying
any layout change before a full-baseline verify.

Usage:
  python tools/metrics/canary_bottom_15.py
"""
import os
import shutil
import sys
import tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_DIR = os.path.join(ROOT, "tools", "golden-test", "documents", "docx")

# Bottom-15 docs by min(SSIM)
TARGETS = [
    "d77a58485f16_20240705_resources_data_outline_08",
    "b837808d0555_20240705_resources_data_guideline_02",
    "29dc6e8943fe_order_01",
    "2ea81a8441cc_0025006-192",
    "e3c545fac7a7_LOD_Handbook",
    "b35123fe8efc_tokumei_08_01",
    "1ec1091177b1_006",
    "04b88e7e0b25_index-19",
    "d4d126dfe1d9_tokumei_08_01-3",
    "34140b9c5662_index-14",
    "459f05f1e877_kyodokenkyuyoushiki01",
    "6514f214e482_tokumei_08_01-2",
    "db9ca18368cd_20241122_resource_open_data_01",
    "1636d28e2c46_tokumei_08_04",
    "a1d6e4efa2e7_tokumei_08_01-4",
]


def main():
    tmp = tempfile.mkdtemp(prefix="oxi_r35_canary_")
    found = 0
    for stem in TARGETS:
        src = os.path.join(DOCX_DIR, stem + ".docx")
        if not os.path.exists(src):
            print(f"# missing: {stem}")
            continue
        dst = os.path.join(tmp, stem + ".docx")
        try:
            os.symlink(src, dst)
        except (OSError, NotImplementedError):
            shutil.copy2(src, dst)
        found += 1
    print(f"# staged {found}/{len(TARGETS)} docs at {tmp}")

    os.environ["OXI_ALLOW_REGRESSION"] = "1"
    sys.path.insert(0, ROOT)
    from pipeline.verify import verify
    ok = verify(tmp, limit=0)
    print(f"\nverify returned: {ok}")
    print(f"(temp dir kept for inspection: {tmp})")


if __name__ == "__main__":
    main()
