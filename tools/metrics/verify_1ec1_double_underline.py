"""Single-doc canary for the DWrite double-underline patch.

Only 1ec1091177b1_006 in the 267-doc baseline uses `<w:u w:val="double"/>`,
so the canary scope is one document. Snapshots ssim_baseline.json before
running pipeline.verify and restores it after, so the canary cannot
accidentally update the baseline.
"""
import os, shutil, sys, tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(ROOT, "tools", "golden-test", "documents", "docx",
                    "1ec1091177b1_006.docx")
BASELINE = os.path.join(ROOT, "pipeline_data", "ssim_baseline.json")


def main():
    oxi_png = os.path.join(ROOT, "pipeline_data", "oxi_png", "1ec1091177b1_006")
    if os.path.isdir(oxi_png):
        shutil.rmtree(oxi_png)
        print("cleared oxi_png cache")

    with open(BASELINE, "rb") as f:
        baseline_snapshot = f.read()

    tmp = tempfile.mkdtemp(prefix="oxi_1ec1_")
    dst = os.path.join(tmp, os.path.basename(DOCX))
    try:
        os.symlink(DOCX, dst)
    except (OSError, NotImplementedError):
        shutil.copy2(DOCX, dst)
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
