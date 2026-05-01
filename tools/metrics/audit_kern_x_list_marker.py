"""依頼 1 拡張: kern × list_marker (numPr) cross-tab.

For each docx:
  - effective kern (Normal style > docDefaults)
  - count of paragraphs with numPr (list_marker)
  - count of paragraphs total
  - ratio of list paragraphs

Then cross-tab against:
  - SSIM distribution
  - R17 winners/losers

Goal: identify why R17's list_marker gate had partial success.
Hypothesis: docs with kern often also have list_marker → R17's gate
correlates with kern by accident.
"""
import os
import json
import zipfile
import re
import sys
from glob import glob
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = os.path.abspath("tools/golden-test/documents/docx")
SSIM_BASELINE = os.path.abspath("pipeline_data/ssim_baseline.json")
RESULT_PATH = os.path.abspath(
    "pipeline_data/kern_x_list_marker_audit_2026-05-02.json")


def extract_kern_and_list(docx_path):
    """Returns: dict with kern info + paragraph counts."""
    out = {
        "dd_kern": None, "normal_kern": None,
        "effective_kern": None, "kern_present": False,
        "n_paragraphs_total": 0,
        "n_paragraphs_with_numPr": 0,
        "n_paragraphs_with_textcontent": 0,
        "list_para_ratio": 0.0,
    }
    try:
        with zipfile.ZipFile(docx_path) as z:
            try:
                styles = z.read("word/styles.xml").decode("utf-8")
            except KeyError:
                return {"error": "styles.xml missing"}
            try:
                doc = z.read("word/document.xml").decode("utf-8")
            except KeyError:
                doc = ""

        # docDefaults rPr
        m_dd = re.search(
            r'<w:docDefaults>.*?<w:rPrDefault>.*?<w:rPr>(.*?)</w:rPr>',
            styles, re.DOTALL)
        if m_dd:
            km = re.search(r'<w:kern\s+w:val="([^"]+)"\s*/?>',
                            m_dd.group(1))
            if km:
                try:
                    out["dd_kern"] = int(km.group(1))
                except ValueError:
                    out["dd_kern"] = km.group(1)

        # Normal style rPr
        m_normal = re.search(
            r'<w:style[^>]*w:type="paragraph"[^>]*w:default="1"[^>]*>(.*?)</w:style>',
            styles, re.DOTALL)
        if not m_normal:
            m_normal = re.search(
                r'<w:style[^>]*w:default="1"[^>]*w:type="paragraph"[^>]*>(.*?)</w:style>',
                styles, re.DOTALL)
        if m_normal:
            inner = m_normal.group(1)
            rpr_m = re.search(r'<w:rPr>(.*?)</w:rPr>', inner, re.DOTALL)
            if rpr_m:
                km = re.search(r'<w:kern\s+w:val="([^"]+)"\s*/?>',
                                rpr_m.group(1))
                if km:
                    try:
                        out["normal_kern"] = int(km.group(1))
                    except ValueError:
                        out["normal_kern"] = km.group(1)

        out["effective_kern"] = (out["normal_kern"]
                                  if out["normal_kern"] is not None
                                  else out["dd_kern"])
        out["kern_present"] = (out["effective_kern"] is not None
                                and out["effective_kern"] != 0)

        # Paragraph count + numPr presence in document.xml
        # Each <w:p>...</w:p> = 1 paragraph
        if doc:
            paras = re.findall(r'<w:p[\s>](.*?)</w:p>', doc, re.DOTALL)
            out["n_paragraphs_total"] = len(paras)
            n_with_num = 0
            n_with_text = 0
            for p in paras:
                if "<w:numPr>" in p:
                    n_with_num += 1
                if "<w:t>" in p or "<w:t " in p:
                    n_with_text += 1
            out["n_paragraphs_with_numPr"] = n_with_num
            out["n_paragraphs_with_textcontent"] = n_with_text
            if n_with_text > 0:
                out["list_para_ratio"] = round(n_with_num / n_with_text, 4)
        return out
    except Exception as e:
        return {"error": str(e)}


def doc_id_from_filename(filename):
    base = os.path.basename(filename).replace(".docx", "")
    m = re.match(r'([0-9a-f]{12,})', base)
    return m.group(1) if m else base


R17_BIG_WINNERS = [
    ("d77a58485f16", 3), ("d77a58485f16", 6),
    ("683ffcab86e2", 2),
    ("0e7af1ae8f21", 6),
    ("3a4f9fbe1a83", 23), ("3a4f9fbe1a83", 60),
]
R17_BIG_LOSERS = [
    ("7f272a2dfd3b", 1),
    ("ed025cbecffb", 1), ("ed025cbecffb", 2),
    ("3a4f9fbe1a83", 64),
]


def main():
    with open(SSIM_BASELINE, encoding="utf-8") as f:
        ssim = json.load(f)
    all_docx = sorted(glob(os.path.join(DOCX_DIR, "*.docx")))
    audit = []
    for path in all_docx:
        info = extract_kern_and_list(path)
        if "error" in info:
            audit.append({"doc_id_full": os.path.basename(path),
                          "error": info["error"]})
            continue
        doc_id_full = os.path.basename(path).replace(".docx", "")
        doc_id_short = doc_id_from_filename(path)
        ssim_for_doc = ssim.get(doc_id_full, {})
        page_ssims = [(k, v) for k, v in ssim_for_doc.items()
                      if isinstance(v, (int, float))]
        min_ssim = (min(s for _, s in page_ssims)
                    if page_ssims else None)
        rec = {
            "doc_id_full": doc_id_full,
            "doc_id_short": doc_id_short,
            "min_ssim": (round(min_ssim, 4) if min_ssim is not None
                         else None),
            "page_ssims": dict(page_ssims),
        }
        rec.update(info)
        audit.append(rec)

    # Cross-tab
    has_kern = [a for a in audit if a.get("kern_present")]
    no_kern = [a for a in audit if a.get("kern_present") is False]
    print(f"Total docs: {len(audit)}")
    print(f"With effective kern: {len(has_kern)}")
    print(f"Without: {len(no_kern)}")

    # 4-way: kern × list_marker
    def has_list(a, threshold=0.05):
        return a.get("list_para_ratio", 0) >= threshold

    cross = Counter()
    for a in audit:
        if "error" in a:
            continue
        k = "kern" if a.get("kern_present") else "no_kern"
        l = ("has_list_para" if has_list(a, 0.05)
             else "no_list_para")
        cross[(k, l)] += 1

    print("\n=== kern × list_marker cross-tab (list_marker_ratio≥5%) ===")
    for (k, l), n in sorted(cross.items()):
        print(f"  {k}, {l}: {n}")

    # Mean SSIM per quadrant
    print("\n=== Mean min_ssim per quadrant ===")
    for k_label in ["kern", "no_kern"]:
        for l_label in ["has_list_para", "no_list_para"]:
            ssims = []
            for a in audit:
                if "error" in a:
                    continue
                if a.get("kern_present") != (k_label == "kern"):
                    continue
                if has_list(a, 0.05) != (l_label == "has_list_para"):
                    continue
                if a["min_ssim"] is not None:
                    ssims.append(a["min_ssim"])
            if ssims:
                ssims.sort()
                print(f"  {k_label}, {l_label}: n={len(ssims)} "
                      f"mean={sum(ssims)/len(ssims):.4f} "
                      f"min={ssims[0]:.4f} max={ssims[-1]:.4f}")
            else:
                print(f"  {k_label}, {l_label}: n=0")

    # R17 categories
    def find(did):
        for a in audit:
            if a.get("doc_id_short") == did:
                return a

    print("\n=== R17 big_winners (kern × list_marker) ===")
    for did, page in R17_BIG_WINNERS:
        a = find(did)
        if a:
            print(f"  {did} p.{page}: "
                  f"kern={a.get('effective_kern')!r} "
                  f"list_ratio={a.get('list_para_ratio'):.3f} "
                  f"({a.get('n_paragraphs_with_numPr')}/"
                  f"{a.get('n_paragraphs_with_textcontent')})")

    print("\n=== R17 big_losers (kern × list_marker) ===")
    for did, page in R17_BIG_LOSERS:
        a = find(did)
        if a:
            print(f"  {did} p.{page}: "
                  f"kern={a.get('effective_kern')!r} "
                  f"list_ratio={a.get('list_para_ratio'):.3f} "
                  f"({a.get('n_paragraphs_with_numPr')}/"
                  f"{a.get('n_paragraphs_with_textcontent')})")

    cross_str = {f"{k}|{l}": n for (k, l), n in cross.items()}
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump({"audit": audit, "cross_tab": cross_str}, f,
                  ensure_ascii=False, indent=2)
    print(f"\nWrote audit to {RESULT_PATH}")


if __name__ == "__main__":
    main()
