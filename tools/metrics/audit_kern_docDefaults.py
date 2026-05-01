"""依頼 1: kern docDefaults audit across baseline docs.

For each docx in tools/golden-test/documents/docx/:
- Extract <w:kern w:val="N"/> from word/styles.xml docDefaults rPrDefault rPr
- Cross-tab against:
  - R17 big_winners (d77a p3/p6, 683f p2, 0e7af1 p6, 3a4f p23/p60)
  - R17 big_losers (7f272a p1, ed025 p1/p2, 3a4f p64)
  - R31 winner (3a4f_p64) / loser (3a4f_p42)
  - All-doc SSIM distribution

Output: pipeline_data/kern_audit_2026-05-02.json + summary stats
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
RESULT_PATH = os.path.abspath("pipeline_data/kern_audit_2026-05-02.json")

# R17 categories per user's classification
R17_BIG_WINNERS = [
    ("d77a58485f16", 3),
    ("d77a58485f16", 6),
    ("683ffcab86e2", 2),
    ("0e7af1ae8f21", 6),
    ("3a4f9fbe1a83", 23),
    ("3a4f9fbe1a83", 60),
]
R17_BIG_LOSERS = [
    ("7f272a2dfd3b", 1),
    ("ed025cbecffb", 1),
    ("ed025cbecffb", 2),
    ("3a4f9fbe1a83", 64),
]
R31_WINNERS = [("3a4f9fbe1a83", 64)]
R31_LOSERS = [("3a4f9fbe1a83", 42)]


def extract_kern(docx_path):
    """Extract w:kern val from docDefaults rPrDefault rPr AND Normal style rPr.
    Returns: dict with dd_kern, normal_kern, effective_kern.

    Resolution priority: run rPr > Normal style rPr > docDefaults rPr.
    For audit, "effective" = first non-None of (Normal, docDefaults).
    """
    try:
        with zipfile.ZipFile(docx_path) as z:
            try:
                styles = z.read("word/styles.xml").decode("utf-8")
            except KeyError:
                return {"error": "styles.xml missing"}
        # docDefaults rPrDefault rPr
        m_dd = re.search(
            r'<w:docDefaults>.*?<w:rPrDefault>.*?<w:rPr>(.*?)</w:rPr>',
            styles, re.DOTALL)
        dd_kern_val = None
        if m_dd:
            km = re.search(r'<w:kern\s+w:val="([^"]+)"\s*/?>', m_dd.group(1))
            if km:
                try:
                    dd_kern_val = int(km.group(1))
                except ValueError:
                    dd_kern_val = km.group(1)

        # Normal style rPr (default paragraph style: w:default="1" w:type="paragraph")
        m_normal = re.search(
            r'<w:style[^>]*w:type="paragraph"[^>]*w:default="1"[^>]*>(.*?)</w:style>',
            styles, re.DOTALL)
        if not m_normal:
            m_normal = re.search(
                r'<w:style[^>]*w:default="1"[^>]*w:type="paragraph"[^>]*>(.*?)</w:style>',
                styles, re.DOTALL)
        normal_kern_val = None
        if m_normal:
            normal_inner = m_normal.group(1)
            # Find kern within Normal's rPr
            rpr_m = re.search(r'<w:rPr>(.*?)</w:rPr>', normal_inner, re.DOTALL)
            if rpr_m:
                km = re.search(r'<w:kern\s+w:val="([^"]+)"\s*/?>',
                                rpr_m.group(1))
                if km:
                    try:
                        normal_kern_val = int(km.group(1))
                    except ValueError:
                        normal_kern_val = km.group(1)
        # Effective kern: Normal overrides docDefaults
        effective = (normal_kern_val if normal_kern_val is not None
                     else dd_kern_val)
        return {
            "dd_kern": dd_kern_val,
            "normal_kern": normal_kern_val,
            "effective_kern": effective,
            "kern_present": effective is not None and effective != 0,
        }
    except Exception as e:
        return {"error": f"{e}"}


def doc_id_from_filename(filename):
    """Extract leading hex prefix (12+ chars) before underscore as doc id."""
    base = os.path.basename(filename).replace(".docx", "")
    m = re.match(r'([0-9a-f]{12,})', base)
    return m.group(1) if m else base


def main():
    # Load SSIM baseline
    with open(SSIM_BASELINE, encoding="utf-8") as f:
        ssim = json.load(f)
    # Audit all docx
    all_docx = sorted(glob(os.path.join(DOCX_DIR, "*.docx")))
    print(f"Auditing {len(all_docx)} docx files...\n", flush=True)
    audit = []
    for path in all_docx:
        kern_info = extract_kern(path)
        doc_id_full = os.path.basename(path).replace(".docx", "")
        doc_id_short = doc_id_from_filename(path)
        # Find SSIM page-min
        ssim_for_doc = ssim.get(doc_id_full, {})
        page_ssims = []
        for k, v in ssim_for_doc.items():
            if isinstance(v, (int, float)):
                page_ssims.append((str(k), v))
        min_ssim = (min(page_ssims, key=lambda t: t[1])[1]
                    if page_ssims else None)
        rec = {
            "doc_id_full": doc_id_full,
            "doc_id_short": doc_id_short,
            "n_pages_ssim": len(page_ssims),
            "min_ssim": (round(min_ssim, 4) if min_ssim is not None
                          else None),
            "page_ssims": dict(page_ssims),
        }
        rec.update(kern_info)
        # Backwards compatible aliases
        rec["kern_val"] = kern_info.get("effective_kern")
        rec["kern_present"] = kern_info.get("kern_present", False)
        audit.append(rec)

    # Stats
    total = len(audit)
    with_kern = [a for a in audit if a["kern_present"]]
    without_kern = [a for a in audit if not a["kern_present"]]
    by_kern_val = Counter(a["kern_val"] for a in with_kern)
    by_source = Counter()
    for a in with_kern:
        if a.get("normal_kern") is not None:
            by_source["Normal style"] += 1
        elif a.get("dd_kern") is not None:
            by_source["docDefaults"] += 1
    print(f"Total docs: {total}")
    print(f"With kern (effective): {len(with_kern)} "
          f"({len(with_kern)*100/total:.1f}%)")
    print(f"Without kern: {len(without_kern)}")
    print(f"Effective kern by val: {sorted(by_kern_val.items())}")
    print(f"Effective kern source: {dict(by_source)}")

    # SSIM distribution
    def ssim_stats(group, label):
        ssims = [a["min_ssim"] for a in group if a["min_ssim"] is not None]
        if not ssims:
            print(f"\n{label}: no SSIM data")
            return
        ssims.sort()
        print(f"\n{label} (n={len(ssims)}):")
        print(f"  min={min(ssims):.4f} max={max(ssims):.4f} "
              f"mean={sum(ssims)/len(ssims):.4f}")
        print(f"  ssim<0.7: {sum(1 for s in ssims if s < 0.7)}")
        print(f"  0.7≤ssim<0.85: {sum(1 for s in ssims if 0.7 <= s < 0.85)}")
        print(f"  ssim≥0.85: {sum(1 for s in ssims if s >= 0.85)}")

    ssim_stats(with_kern, "WITH kern (min_ssim distribution)")
    ssim_stats(without_kern, "WITHOUT kern (min_ssim distribution)")

    # Cross-tab against R17/R31 categories
    def find_audit(doc_id_short):
        for a in audit:
            if a["doc_id_short"] == doc_id_short:
                return a
        return None

    def fmt_kern(a):
        return (f"dd={a.get('dd_kern')!r} normal={a.get('normal_kern')!r} "
                f"eff={a.get('effective_kern')!r}")

    print("\n=== R17 big_winners (kern dd / normal / effective) ===")
    for did, page in R17_BIG_WINNERS:
        a = find_audit(did)
        if a:
            page_ssim = a["page_ssims"].get(str(page))
            print(f"  {did} p.{page}: {fmt_kern(a)} ssim={page_ssim}")
        else:
            print(f"  {did}: NOT FOUND")

    print("\n=== R17 big_losers (kern dd / normal / effective) ===")
    for did, page in R17_BIG_LOSERS:
        a = find_audit(did)
        if a:
            page_ssim = a["page_ssims"].get(str(page))
            print(f"  {did} p.{page}: {fmt_kern(a)} ssim={page_ssim}")
        else:
            print(f"  {did}: NOT FOUND")

    print("\n=== R31 winner ===")
    for did, page in R31_WINNERS:
        a = find_audit(did)
        if a:
            page_ssim = a["page_ssims"].get(str(page))
            print(f"  {did} p.{page}: {fmt_kern(a)} ssim={page_ssim}")

    print("\n=== R31 loser ===")
    for did, page in R31_LOSERS:
        a = find_audit(did)
        if a:
            page_ssim = a["page_ssims"].get(str(page))
            print(f"  {did} p.{page}: {fmt_kern(a)} ssim={page_ssim}")

    # Save
    out = {
        "summary": {
            "total": total,
            "with_kern": len(with_kern),
            "without_kern": len(without_kern),
            "kern_val_distribution": dict(by_kern_val),
        },
        "audit": audit,
        "r17_big_winners": [
            {"doc": d, "page": p,
             "audit": find_audit(d)} for d, p in R17_BIG_WINNERS
        ],
        "r17_big_losers": [
            {"doc": d, "page": p,
             "audit": find_audit(d)} for d, p in R17_BIG_LOSERS
        ],
        "r31_winners": [
            {"doc": d, "page": p,
             "audit": find_audit(d)} for d, p in R31_WINNERS
        ],
        "r31_losers": [
            {"doc": d, "page": p,
             "audit": find_audit(d)} for d, p in R31_LOSERS
        ],
    }
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote audit to {RESULT_PATH}")


if __name__ == "__main__":
    main()
