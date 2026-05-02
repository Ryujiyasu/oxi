"""
Delete pipeline_data/oxi_png/<doc_id>/ for every baseline doc whose docx
contains <w:wordWrap w:val="off"/> (or "0"). Forces verify to re-render.
"""
import os
import re
import shutil
import sys
import zipfile

DOCX_DIR = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..", "tools", "golden-test", "documents", "docx"
))
OXI_PNG_DIR = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..", "pipeline_data", "oxi_png"
))

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
            affected.append(fn[:-5])  # strip .docx
    print(f"# affected docs: {len(affected)}")
    cleared = 0
    for stem in affected:
        d = os.path.join(OXI_PNG_DIR, stem)
        if os.path.isdir(d):
            shutil.rmtree(d)
            cleared += 1
            print(f"  cleared: {stem}")
        else:
            print(f"  (no cache) {stem}")
    print(f"\n# cleared {cleared}/{len(affected)} cache dirs")


if __name__ == "__main__":
    main()
