"""Quick SSIM canary on the 3 tokumei docs that S109 \\n-fix affects.

Snapshots ssim_baseline.json before running pipeline.verify and restores
it after.
"""
import os, shutil, sys, tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCS = [
    "a1d6e4efa2e7_tokumei_08_01-4.docx",
    "d4d126dfe1d9_tokumei_08_01-3.docx",
    "de6e32b5960b_tokumei_08_01-1.docx",
]
BASELINE = os.path.join(ROOT, "pipeline_data", "ssim_baseline.json")


def main():
    for fname in DOCS:
        doc_id = fname.replace(".docx", "")
        oxi_png = os.path.join(ROOT, "pipeline_data", "oxi_png", doc_id)
        if os.path.isdir(oxi_png):
            shutil.rmtree(oxi_png)

    with open(BASELINE, "rb") as f:
        baseline_snapshot = f.read()

    tmp = tempfile.mkdtemp(prefix="oxi_softbreak_")
    for fname in DOCS:
        src = os.path.join(ROOT, "tools", "golden-test", "documents", "docx", fname)
        dst = os.path.join(tmp, fname)
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
            f.write(baseline_snapshot)
        print("# baseline restored from snapshot")


if __name__ == "__main__":
    main()
