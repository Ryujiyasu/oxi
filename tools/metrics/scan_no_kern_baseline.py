"""
Scan pipeline_data/docx/ for baseline docs WITHOUT <w:kern> in docDefaults.

Output: list of candidate docs that:
  - Have NO w:kern in word/styles.xml docDefaults rPr
  - Contain yakumono characters in body text (so we have something to measure)
  - Are NOT synthetic test fixtures (heuristic: filename matches typical baseline pattern)

Output: pipeline_data/no_kern_candidates.json
"""
import os
import zipfile
import json
import re
import xml.etree.ElementTree as ET

DOCX_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data", "docx")
OUT_JSON = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data", "no_kern_candidates.json")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

# Yakumono characters (subset for detection — JIS X 4051 type A/B/C)
YAKUMONO_CHARS = set(
    "（「『【〔｛〈《［"          # Type A: opening
    "）」』】〕｝〉》］、。，．"      # Type B: closing + period
    "・：；！？ー―／＼"          # Type C: middle/separator
    "“”‘’"                          # quotes
)


def has_kern_in_docDefaults(zip_obj, names):
    """Return True if styles.xml has <w:kern> in docDefaults/rPrDefault."""
    if "word/styles.xml" not in names:
        return False
    try:
        with zip_obj.open("word/styles.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()
        # docDefaults/rPrDefault/rPr/kern
        for kern in root.findall(".//w:docDefaults//w:rPrDefault//w:rPr/w:kern", NS):
            v = kern.get(f"{{{NS['w']}}}val")
            if v and int(v) > 0:
                return True
        return False
    except Exception:
        return False


def has_kern_anywhere(zip_obj, names):
    """Check if ANY w:kern element exists (in any style or run)."""
    if "word/styles.xml" not in names:
        return False
    try:
        with zip_obj.open("word/styles.xml") as f:
            data = f.read().decode("utf-8")
        return "<w:kern" in data or "<w:kern " in data
    except Exception:
        return False


def has_run_kern_in_body(zip_obj, names):
    """Check if document.xml has any <w:kern> in any run."""
    if "word/document.xml" not in names:
        return False
    try:
        with zip_obj.open("word/document.xml") as f:
            data = f.read().decode("utf-8")
        return "<w:kern" in data
    except Exception:
        return False


def count_yakumono_in_body(zip_obj, names):
    """Count yakumono characters in document.xml body text."""
    if "word/document.xml" not in names:
        return 0
    try:
        with zip_obj.open("word/document.xml") as f:
            data = f.read().decode("utf-8")
        # Extract text inside <w:t>...</w:t>
        body_text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", data))
        return sum(1 for ch in body_text if ch in YAKUMONO_CHARS)
    except Exception:
        return 0


def count_consecutive_yakumono_pairs(zip_obj, names):
    """Count consecutive yakumono pairs in body text (= compression candidates)."""
    if "word/document.xml" not in names:
        return 0
    try:
        with zip_obj.open("word/document.xml") as f:
            data = f.read().decode("utf-8")
        body_text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", data))
        # Decode XML entities
        body_text = body_text.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
        n = 0
        for i in range(len(body_text) - 1):
            if body_text[i] in YAKUMONO_CHARS and body_text[i+1] in YAKUMONO_CHARS:
                n += 1
        return n
    except Exception:
        return 0


def scan_doc(path):
    info = {"path": os.path.basename(path)}
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            info["kern_docDefaults"] = has_kern_in_docDefaults(z, names)
            info["kern_anywhere_in_styles"] = has_kern_anywhere(z, names)
            info["kern_in_body_runs"] = has_run_kern_in_body(z, names)
            info["yakumono_count_in_body"] = count_yakumono_in_body(z, names)
            info["consecutive_pairs_in_body"] = count_consecutive_yakumono_pairs(z, names)
    except Exception as e:
        info["error"] = str(e)
    return info


def main():
    if not os.path.isdir(DOCX_DIR):
        print(f"DOCX_DIR not found: {DOCX_DIR}")
        return

    candidates = []
    all_results = []
    for fname in sorted(os.listdir(DOCX_DIR)):
        if not fname.endswith(".docx"):
            continue
        path = os.path.join(DOCX_DIR, fname)
        info = scan_doc(path)
        all_results.append(info)
        # Candidate criteria:
        #   - No kern in docDefaults
        #   - No kern in body runs (not just defaults)
        #   - Has yakumono characters
        #   - Skip synthetic test fixtures (RUBY_*, VW_V*, etc.)
        is_synthetic = bool(re.match(r"^(RUBY_|VW_V|FT_|FE_|SP_|SR_|CT_|TR_|CG_|RT_|HR_|BL_|test_|comments_|section_break_|header_)", fname))
        if (not info.get("kern_docDefaults", False)
            and not info.get("kern_in_body_runs", False)
            and info.get("yakumono_count_in_body", 0) >= 5
            and not is_synthetic):
            candidates.append(info)

    # Sort by consecutive pairs (the actual compression candidates)
    candidates.sort(key=lambda x: -x.get("consecutive_pairs_in_body", 0))

    print(f"Total docx scanned: {len(all_results)}")
    print(f"No-kern candidates with yakumono: {len(candidates)}")
    print("\nTop 20 candidates by CONSECUTIVE pair count:")
    for c in candidates[:20]:
        print(f"  {c['path']:50s} yakumono={c.get('yakumono_count_in_body'):4} "
              f"pairs={c.get('consecutive_pairs_in_body'):3} "
              f"kernStyles={c.get('kern_anywhere_in_styles'):>5}")

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump({"all": all_results, "candidates": candidates}, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {OUT_JSON}")


if __name__ == "__main__":
    main()
