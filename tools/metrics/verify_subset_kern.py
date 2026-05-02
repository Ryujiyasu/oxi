"""
Canary verify: stage all baseline docs that have effective kern (62 docs)
into a temp dir and run pipeline.verify against them. Reports per-page
diffs but does NOT fail on regressions (OXI_ALLOW_REGRESSION=1).
"""
import json
import os
import shutil
import sys
import tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_DIR = os.path.join(ROOT, "tools", "golden-test", "documents", "docx")
AUDIT = os.path.join(ROOT, "pipeline_data", "kern_audit_2026-05-02.json")


def main():
    with open(AUDIT, "r", encoding="utf-8") as f:
        raw = json.load(f)
    audit = raw["audit"]
    affected = []
    for r in audit:
        if r.get("kern_present"):
            stem = r.get("doc_id_full")
            if stem:
                src = os.path.join(DOCX_DIR, stem + ".docx")
                if os.path.exists(src):
                    affected.append(stem + ".docx")
    print(f"# canary docs: {len(affected)}")

    # Clear any cached oxi_png for these (force re-render).
    oxi_png = os.path.join(ROOT, "pipeline_data", "oxi_png")
    cleared = 0
    for fn in affected:
        d = os.path.join(oxi_png, fn[:-5])
        if os.path.isdir(d):
            shutil.rmtree(d)
            cleared += 1
    print(f"# cleared {cleared} oxi_png cache dirs")

    tmp = tempfile.mkdtemp(prefix="oxi_kern_canary_")
    for fn in affected:
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
    ok = verify(tmp, limit=0)
    print(f"\nverify returned: {ok}")
    print(f"(temp dir kept for inspection: {tmp})")


if __name__ == "__main__":
    main()
