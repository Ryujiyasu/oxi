"""Two-doc canary: d77a + ed025 — used for the 82de3fa revert trade-off.

Validates that reverting the trailing-U+3000 immune-from-wrap rule
restores d77a p.8/9/10 (which had regressed -0.099 net) while losing
ed025 p.1's +0.042 win — net +0.057 on the bottom-bucket pair.

Snapshots ssim_baseline.json before/after the verify call to prevent
the auto-update from corrupting the canonical baseline values.
"""
import os, shutil, sys, tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_DIR = os.path.join(ROOT, "tools", "golden-test", "documents", "docx")
BASELINE = os.path.join(ROOT, "pipeline_data", "ssim_baseline.json")

DOCS = [
    "d77a58485f16_20240705_resources_data_outline_08.docx",
    "ed025cbecffb_index-23.docx",
]


def main():
    oxi_png = os.path.join(ROOT, "pipeline_data", "oxi_png")
    for fn in DOCS:
        d = os.path.join(oxi_png, fn[:-5])
        if os.path.isdir(d):
            shutil.rmtree(d)
            print(f"cleared cache: {fn[:-5]}")

    with open(BASELINE, "rb") as f:
        snap = f.read()

    tmp = tempfile.mkdtemp(prefix="oxi_d77a_ed025_")
    for fn in DOCS:
        src = os.path.join(DOCX_DIR, fn)
        dst = os.path.join(tmp, fn)
        try:
            os.symlink(src, dst)
        except (OSError, NotImplementedError):
            shutil.copy2(src, dst)
    print(f"# staged at {tmp}")

    os.environ["OXI_ALLOW_REGRESSION"] = "1"
    sys.path.insert(0, ROOT)
    from pipeline.verify import verify
    try:
        verify(tmp, limit=0)
    finally:
        with open(BASELINE, "wb") as f:
            f.write(snap)
        print("# baseline restored from snapshot")


if __name__ == "__main__":
    main()
