"""
Canary verify: run pipeline.verify against the 34 wordWrap=off docs only.
Stages affected .docx into a temp dir and points verify at it.

Reports per-page diffs but does NOT update the baseline (we want to inspect
before promoting). Sets OXI_ALLOW_REGRESSION=1 so verify reports rather
than failing on any single regression.
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_DIR = os.path.join(ROOT, "tools", "golden-test", "documents", "docx")

PAT = re.compile(rb'<w:wordWrap\s+w:val="(off|0)"\s*/?>')


def has_wordwrap_off(path):
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                if name.endswith(".xml"):
                    try:
                        if PAT.search(zf.read(name)):
                            return True
                    except Exception:
                        continue
        return False
    except zipfile.BadZipFile:
        return False


def main():
    affected = []
    for fn in sorted(os.listdir(DOCX_DIR)):
        if fn.endswith(".docx") and has_wordwrap_off(os.path.join(DOCX_DIR, fn)):
            affected.append(fn)
    print(f"# canary docs: {len(affected)}")

    # Stage symlinks (copies on Windows fallback) into temp
    tmp = tempfile.mkdtemp(prefix="oxi_canary_")
    for fn in affected:
        src = os.path.join(DOCX_DIR, fn)
        dst = os.path.join(tmp, fn)
        try:
            os.symlink(src, dst)
        except (OSError, NotImplementedError):
            shutil.copy2(src, dst)
    print(f"# staged at {tmp}")

    # Run verify with allow-regression so it reports rather than failing.
    os.environ["OXI_ALLOW_REGRESSION"] = "1"
    sys.path.insert(0, ROOT)
    from pipeline.verify import verify
    ok = verify(tmp, limit=0)
    print(f"\nverify returned: {ok}")
    print(f"(temp dir kept for inspection: {tmp})")


if __name__ == "__main__":
    main()
