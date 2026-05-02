"""Y0 floating-table refined formula canary.

Tests §19.10 Y0 fixes on the 5 baseline docs that have text-anchored
floating tables (4 of them in current bottom-15: 2ea81a / 1ec1 / 3a4f9f /
ed025 / 459f / e201249db062).

History:
  2026-05-03 attempt: +0.5pt floor margin alone — bottom-5 floor regressed
    (2ea81a p.2 -0.0011). Reverted. Refined formula needs cells-based
    Y0 = (alloc_anchor + natural_anchor)/2 + 0.5pt — flat +0.5pt is wrong
    for tall (2-cell) anchors.
"""
import os
import shutil
import sys
import tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_DIR = os.path.join(ROOT, "tools", "golden-test", "documents", "docx")

TARGETS = [
    "2ea81a8441cc_0025006-192",
    "1ec1091177b1_006",
    "3a4f9fbe1a83_001620506",
    "ed025cbecffb_index-23",
    "459f05f1e877_kyodokenkyuyoushiki01",
    "e201249db062_tokumei_08_05",
]


def main():
    tmp = tempfile.mkdtemp(prefix="oxi_y0_canary_")
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


if __name__ == "__main__":
    main()
