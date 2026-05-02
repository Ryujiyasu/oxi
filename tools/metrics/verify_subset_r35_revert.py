"""
Re-canary the 5 docs that regressed worst in the R35 absorb-budget attempt
to confirm the revert restored pre-fix behavior. Clears their oxi_png cache
(they're stale from the broken render).
"""
import os
import shutil
import sys
import tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_DIR = os.path.join(ROOT, "tools", "golden-test", "documents", "docx")

DOCS = [
    "d77a58485f16_20240705_resources_data_outline_08.docx",
    "c7b923e5c616_20240705_resources_data_outline_06.docx",
    "ed025cbecffb_index-23.docx",
    "a5ccbe425525_kyodokenkyuyoushiki05.docx",
    "3a4f9fbe1a83_001620506.docx",
]


def main():
    oxi_png = os.path.join(ROOT, "pipeline_data", "oxi_png")
    for fn in DOCS:
        d = os.path.join(oxi_png, fn[:-5])
        if os.path.isdir(d):
            shutil.rmtree(d)
            print(f"  cleared cache: {fn[:-5]}")

    tmp = tempfile.mkdtemp(prefix="oxi_revert_canary_")
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
    verify(tmp, limit=0)


if __name__ == "__main__":
    main()
