"""Scan baseline docs for OOXML features.

Detects:
- Ruby (w:ruby) — furigana
- Vertical writing (w:textDirection, lr-tb-v, tb-rl-v)
- OMML math (m:oMath)
- Drop caps (w:framePr w:dropCap)
- Numbering (w:numPr)
- Table styles (w:tblStyle)
- Field codes (w:fldChar)
- Comments (w:commentReference)

Output: doc_id → list of feature counts.
"""
import os, re, sys, zipfile
from pathlib import Path
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
DOCS_DIR = REPO_ROOT / "tools" / "golden-test" / "documents" / "docx"


def scan_doc(docx_path):
    counts = defaultdict(int)
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("word/document.xml").decode("utf-8", errors="replace")
    except Exception:
        return counts
    counts["ruby"] = len(re.findall(r"<w:ruby\b", xml))
    counts["rubyBase"] = len(re.findall(r"<w:rubyBase\b", xml))
    counts["rt"] = len(re.findall(r"<w:rt\b", xml))
    counts["textDirection"] = len(re.findall(r"<w:textDirection\b", xml))
    counts["vertical"] = sum(
        1 for m in re.finditer(r'<w:textDirection w:val="([^"]+)"', xml)
        if "tb" in m.group(1)
    )
    counts["oMath"] = len(re.findall(r"<m:oMath\b", xml))
    counts["oMathPara"] = len(re.findall(r"<m:oMathPara\b", xml))
    counts["framePr"] = len(re.findall(r"<w:framePr\b", xml))
    counts["dropCap"] = sum(1 for m in re.finditer(r'<w:framePr [^/]*w:dropCap=', xml))
    counts["numPr"] = len(re.findall(r"<w:numPr\b", xml))
    counts["fldChar"] = len(re.findall(r"<w:fldChar\b", xml))
    counts["instrText"] = len(re.findall(r"<w:instrText\b", xml))
    counts["commentRef"] = len(re.findall(r"<w:commentReference\b", xml))
    counts["hyperlink"] = len(re.findall(r"<w:hyperlink\b", xml))
    counts["drawing"] = len(re.findall(r"<w:drawing\b", xml))
    counts["pict"] = len(re.findall(r"<w:pict\b", xml))
    counts["textbox"] = len(re.findall(r"<w:txbxContent\b", xml))
    counts["smartTag"] = len(re.findall(r"<w:smartTag\b", xml))
    counts["bookmark"] = len(re.findall(r"<w:bookmarkStart\b", xml))
    return counts


def main():
    docs = sorted(DOCS_DIR.glob("*.docx"))
    # Aggregate
    features_by_doc = {}
    for fp in docs:
        if fp.name.startswith("~$"):
            continue
        doc_id = fp.name.split("_")[0]
        features_by_doc[doc_id] = scan_doc(fp)

    # Feature totals
    feature_totals = defaultdict(int)
    feature_docs = defaultdict(list)
    for doc_id, counts in features_by_doc.items():
        for feat, n in counts.items():
            if n > 0:
                feature_totals[feat] += n
                feature_docs[feat].append((doc_id, n))

    # Sort by feature priority (focus on layout-impactful)
    priority_features = ["ruby", "rubyBase", "rt", "vertical", "textDirection",
                         "oMath", "oMathPara", "framePr", "dropCap", "numPr",
                         "fldChar", "instrText", "commentRef", "drawing", "pict",
                         "textbox", "hyperlink", "bookmark"]

    print(f"OOXML feature scan ({len(features_by_doc)} docs)")
    print(f"{'feature':<15} {'total_occ':>10} {'n_docs':>7}")
    print("-" * 40)
    for feat in priority_features:
        if feature_totals[feat] > 0:
            print(f"{feat:<15} {feature_totals[feat]:>10} {len(feature_docs[feat]):>7}")

    print()
    # Highlight ruby + vertical + omml docs
    for feat in ["ruby", "rubyBase", "vertical", "oMath", "oMathPara"]:
        if feature_docs[feat]:
            print(f"\nDocs with {feat}:")
            for doc_id, n in sorted(feature_docs[feat], key=lambda x: -x[1])[:10]:
                print(f"  {doc_id:<14} count={n}")


if __name__ == "__main__":
    main()
