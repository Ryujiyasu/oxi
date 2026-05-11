"""Three-doc canary: d77a + ed025 + 1ec1.

Validates the post-merge state (corner_inset + 82de3fa revert):
- d77a: should match baseline (revert restored)
- 1ec1: p.1 should improve +0.031 (corner_inset applied)
- ed025: p.1 should regress -0.042 (82de3fa win lost)
"""
import os, shutil, sys, tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_DIR = os.path.join(ROOT, "tools", "golden-test", "documents", "docx")
BASELINE = os.path.join(ROOT, "pipeline_data", "ssim_baseline.json")

DOCS = [
    "d77a58485f16_20240705_resources_data_outline_08.docx",
    "ed025cbecffb_index-23.docx",
    "1ec1091177b1_006.docx",
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

    tmp = tempfile.mkdtemp(prefix="oxi_3doc_")
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
