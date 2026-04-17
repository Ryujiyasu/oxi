"""OMML inventory across the 177-doc baseline.

Scans every docx for Office Math Markup Language (OMML) usage:
- <m:oMath> (inline math)
- <m:oMathPara> (display math paragraph)
- Tag subtype counts (fraction, sup/sub, sqrt, nary, matrix, etc.)

Output: tools/metrics/output/omml_inventory.json
Fields per doc:
  path, n_oMath, n_oMathPara, tag_counts{tag_name: n}, first_math_snippet

Summary: which docs use OMML, which tags are most common, prioritization
for implementation (frac/sup/sub/sqrt first, matrix/nary later).
"""
import json, re, sys, zipfile
from collections import Counter, defaultdict
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = Path(__file__).resolve().parent.parent.parent / "tools" / "golden-test" / "documents" / "docx"
OUT = Path(__file__).with_name("output") / "omml_inventory.json"
OUT.parent.mkdir(parents=True, exist_ok=True)

# OMML tag names (w:m namespace). We track these explicitly; others bucketed as "other".
OMML_TAGS = [
    "oMath", "oMathPara",            # containers
    "f", "num", "den",                # fraction (f = <m:f>, num = numerator, den = denominator)
    "sSup", "sSub", "sSubSup", "sPre",# sub/superscript (sPre = pre-subscript/superscript)
    "rad", "deg", "e",                # radical (sqrt, nth root)
    "nary",                           # n-ary operator (sum, integral, product)
    "m", "mr",                        # matrix, matrix row
    "acc", "bar", "box", "borderBox", # accent, bar, box, bordered box
    "limLow", "limUpp",               # lower/upper limit
    "func",                           # function apply (sin, cos, etc.)
    "d",                              # delimiter (brackets)
    "groupChr",                       # group character
    "phant",                          # phantom
    "eqArr",                          # equation array
    "r", "t",                         # run, text (leaf nodes)
    "mPr", "rPr",                     # properties containers
    "ctrlPr",                         # control properties
]


def scan_doc(docx_path: Path):
    """Return a dict of OMML usage stats for one docx."""
    try:
        with zipfile.ZipFile(docx_path) as z:
            # Scan all XML parts (document.xml, headers, footers, endnotes, footnotes)
            xml_parts = [n for n in z.namelist() if n.endswith(".xml") and "word/" in n]
            combined_xml = ""
            for name in xml_parts:
                try:
                    combined_xml += z.read(name).decode("utf-8", errors="replace")
                except Exception:
                    continue
    except Exception as e:
        return {"error": str(e)}

    # Regex count OMML tags. Format: <m:tag (...)> or <m:tag/>
    n_oMath = len(re.findall(r"<m:oMath[\s/>]", combined_xml))
    n_oMathPara = len(re.findall(r"<m:oMathPara[\s/>]", combined_xml))

    if n_oMath == 0 and n_oMathPara == 0:
        return {"n_oMath": 0, "n_oMathPara": 0, "tag_counts": {}}

    tag_counts = Counter()
    for tag in OMML_TAGS:
        matches = re.findall(rf"<m:{tag}[\s/>]", combined_xml)
        if matches:
            tag_counts[tag] = len(matches)

    # Other m: tags not in our list
    all_m_tags = re.findall(r"<m:(\w+)[\s/>]", combined_xml)
    other_tags = Counter(t for t in all_m_tags if t not in OMML_TAGS)

    # First OMML snippet (for manual inspection)
    first_match = re.search(r"<m:oMath(?:Para)?[\s/>].{0,300}", combined_xml, re.DOTALL)
    first_snippet = first_match.group()[:300] if first_match else None

    return {
        "n_oMath": n_oMath,
        "n_oMathPara": n_oMathPara,
        "tag_counts": dict(tag_counts),
        "other_m_tags": dict(other_tags),
        "first_snippet": first_snippet,
    }


def main():
    docx_files = sorted(DOCX_DIR.glob("*.docx"))
    print(f"Scanning {len(docx_files)} docx files...")

    results = []
    docs_with_omml = 0
    total_oMath = 0
    total_oMathPara = 0
    global_tag_counts = Counter()

    for docx in docx_files:
        r = scan_doc(docx)
        r["doc_id"] = docx.stem
        results.append(r)
        if r.get("n_oMath", 0) > 0 or r.get("n_oMathPara", 0) > 0:
            docs_with_omml += 1
            total_oMath += r.get("n_oMath", 0)
            total_oMathPara += r.get("n_oMathPara", 0)
            for tag, cnt in r.get("tag_counts", {}).items():
                global_tag_counts[tag] += cnt
            print(f"  {docx.stem[:55]:<55}  oMath={r['n_oMath']:>4}  oMathPara={r['n_oMathPara']:>3}")

    summary = {
        "total_docs": len(docx_files),
        "docs_with_omml": docs_with_omml,
        "total_oMath_occurrences": total_oMath,
        "total_oMathPara_occurrences": total_oMathPara,
        "global_tag_counts": dict(global_tag_counts.most_common()),
        "per_doc_results": results,
    }

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print(f"\n=== SUMMARY ===")
    print(f"Total docs: {summary['total_docs']}")
    print(f"Docs with OMML: {docs_with_omml}")
    print(f"Total <m:oMath>: {total_oMath}")
    print(f"Total <m:oMathPara>: {total_oMathPara}")
    print(f"\nTop 15 OMML tags (by frequency):")
    for tag, cnt in global_tag_counts.most_common(15):
        print(f"  {tag:<15} {cnt:>5}")
    print(f"\nSaved → {OUT}")


if __name__ == "__main__":
    main()
