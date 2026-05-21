"""Identify which docs use specific OOXML features and their current IoU."""
import json, os, re, sys, zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
DOCS_DIR = REPO_ROOT / "tools" / "golden-test" / "documents" / "docx"
IOU_SUMMARY = REPO_ROOT / "pipeline_data" / "element_iou_diff" / "_summary.json"


def main():
    if not IOU_SUMMARY.exists():
        print("No IoU summary")
        return
    with open(IOU_SUMMARY, encoding="utf-8") as f:
        iou_data = json.load(f)
    iou_map = {d["doc_id"]: d["mean_iou"] for d in iou_data["docs"]}

    docs_by_feature = {
        "numPr": [],
        "textbox": [],
        "drawing": [],
        "fldChar": [],
    }
    for fp in sorted(DOCS_DIR.glob("*.docx")):
        if fp.name.startswith("~$"):
            continue
        doc_id = fp.name.split("_")[0]
        try:
            with zipfile.ZipFile(fp) as z:
                xml = z.read("word/document.xml").decode("utf-8", errors="replace")
        except Exception:
            continue
        for feat, regex in [
            ("numPr", r"<w:numPr\b"),
            ("textbox", r"<w:txbxContent\b"),
            ("drawing", r"<w:drawing\b"),
            ("fldChar", r"<w:fldChar\b"),
        ]:
            count = len(re.findall(regex, xml))
            if count > 0:
                docs_by_feature[feat].append((doc_id, count, iou_map.get(doc_id, None)))

    for feat in ["numPr", "textbox", "drawing", "fldChar"]:
        docs = docs_by_feature[feat]
        if not docs: continue
        print(f"\n=== {feat} ({len(docs)} docs) ===")
        print(f"  {'doc_id':<14} {'count':>6} {'iou':>8}")
        for doc_id, count, iou in sorted(docs, key=lambda x: x[2] or 1.0):
            iou_str = f"{iou:.4f}" if iou is not None else "N/A"
            print(f"  {doc_id:<14} {count:>6} {iou_str:>8}")


if __name__ == "__main__":
    main()
