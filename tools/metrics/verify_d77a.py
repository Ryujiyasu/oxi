"""Single-doc canary: render d77a only, report SSIM per page.

Snapshots ssim_baseline.json before running pipeline.verify and restores
it after so the canary cannot accidentally update the baseline file
(verify auto-updates on the zero-regression / improvements-only path).
"""
import os, shutil, sys, tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(ROOT, "tools", "golden-test", "documents", "docx",
                    "d77a58485f16_20240705_resources_data_outline_08.docx")
BASELINE = os.path.join(ROOT, "pipeline_data", "ssim_baseline.json")


def main():
    oxi_png = os.path.join(ROOT, "pipeline_data", "oxi_png",
                           "d77a58485f16_20240705_resources_data_outline_08")
    if os.path.isdir(oxi_png):
        shutil.rmtree(oxi_png)
        print("cleared oxi_png cache")

    with open(BASELINE, "rb") as f:
        baseline_snapshot = f.read()

    tmp = tempfile.mkdtemp(prefix="oxi_d77a_")
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
