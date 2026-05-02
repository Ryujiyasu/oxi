"""Audit characterSpacingControl across baseline docs.

For each docx:
- Extract <w:characterSpacingControl w:val="..."/> from settings.xml
- Cross-tab with kern presence + min_ssim
- Identify how many baseline docs would be affected by Mech 3 gate
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
    "pipeline_data/charSpacingControl_audit_2026-05-02.json")


def extract_charSpacing(docx_path):
    try:
        with zipfile.ZipFile(docx_path) as z:
            try:
                settings = z.read("word/settings.xml").decode("utf-8")
            except KeyError:
                return None
        m = re.search(
            r'<w:characterSpacingControl\s+w:val="([^"]+)"\s*/?>',
            settings)
        return m.group(1) if m else None
    except Exception:
        return None


def extract_kern(docx_path):
    """Effective kern: Normal style > docDefaults."""
    try:
        with zipfile.ZipFile(docx_path) as z:
            try:
                styles = z.read("word/styles.xml").decode("utf-8")
            except KeyError:
                return None
        # docDefaults
        m_dd = re.search(
            r'<w:docDefaults>.*?<w:rPrDefault>.*?<w:rPr>(.*?)</w:rPr>',
            styles, re.DOTALL)
        dd_kern = None
        if m_dd:
            km = re.search(r'<w:kern\s+w:val="([^"]+)"\s*/?>',
                            m_dd.group(1))
            if km:
                try:
                    dd_kern = int(km.group(1))
                except ValueError:
                    dd_kern = km.group(1)
        # Normal style
        m_normal = re.search(
            r'<w:style[^>]*w:type="paragraph"[^>]*w:default="1"[^>]*>(.*?)</w:style>',
            styles, re.DOTALL)
        if not m_normal:
            m_normal = re.search(
                r'<w:style[^>]*w:default="1"[^>]*w:type="paragraph"[^>]*>(.*?)</w:style>',
                styles, re.DOTALL)
        normal_kern = None
        if m_normal:
            inner = m_normal.group(1)
            rpr_m = re.search(r'<w:rPr>(.*?)</w:rPr>', inner, re.DOTALL)
            if rpr_m:
                km = re.search(r'<w:kern\s+w:val="([^"]+)"\s*/?>',
                                rpr_m.group(1))
                if km:
                    try:
                        normal_kern = int(km.group(1))
                    except ValueError:
                        normal_kern = km.group(1)
        return normal_kern if normal_kern is not None else dd_kern
    except Exception:
        return None


def main():
    with open(SSIM_BASELINE, encoding="utf-8") as f:
        ssim = json.load(f)
    all_docx = sorted(glob(os.path.join(DOCX_DIR, "*.docx")))
    audit = []
    for path in all_docx:
        cs = extract_charSpacing(path)
        kern = extract_kern(path)
        doc_id = os.path.basename(path).replace(".docx", "")
        ssim_for_doc = ssim.get(doc_id, {})
        page_ssims = [v for v in ssim_for_doc.values()
                      if isinstance(v, (int, float))]
        min_ssim = min(page_ssims) if page_ssims else None
        audit.append({
            "doc_id": doc_id,
            "charSpacingControl": cs,
            "effective_kern": kern,
            "min_ssim": (round(min_ssim, 4) if min_ssim is not None
                          else None),
        })

    # Stats
    total = len(audit)
    cs_dist = Counter(a["charSpacingControl"] for a in audit)
    print(f"Total docs: {total}")
    print(f"\ncharacterSpacingControl distribution:")
    for k, v in sorted(cs_dist.items(), key=lambda t: (-t[1], str(t[0]))):
        print(f"  {k!r}: {v}")

    # Cross-tab kern × charSpacingControl
    print(f"\nkern × charSpacingControl cross-tab:")
    cross = Counter()
    for a in audit:
        kern_active = (a["effective_kern"] is not None
                        and a["effective_kern"] != 0)
        kern_label = "kern" if kern_active else "no_kern"
        cs = a["charSpacingControl"] or "<absent>"
        cross[(kern_label, cs)] += 1
    for (kl, cs), n in sorted(cross.items()):
        print(f"  {kl}, {cs}: {n}")

    # Mean SSIM per quadrant of (kern × compressPunctuation)
    print(f"\nMean min_ssim per quadrant:")
    for kern_active_label in ["kern", "no_kern"]:
        for cs_compress_label in ["compressPunctuation", "doNotCompress",
                                    "<absent>"]:
            ssims = []
            for a in audit:
                kact = (a["effective_kern"] is not None
                         and a["effective_kern"] != 0)
                if (kact != (kern_active_label == "kern")):
                    continue
                cs = a["charSpacingControl"] or "<absent>"
                if cs != cs_compress_label:
                    continue
                if a["min_ssim"] is not None:
                    ssims.append(a["min_ssim"])
            if ssims:
                print(f"  {kern_active_label}, {cs_compress_label}: "
                      f"n={len(ssims)} "
                      f"mean={sum(ssims)/len(ssims):.4f} "
                      f"min={min(ssims):.4f} max={max(ssims):.4f}")

    # Critical: docs with kern + compressPunctuation are Mech 2/Mech 3 candidates
    candidates = [a for a in audit
                   if (a["effective_kern"] is not None
                       and a["effective_kern"] != 0
                       and a["charSpacingControl"] == "compressPunctuation")]
    print(f"\nMech 2/Mech 3 candidates (kern + compressPunctuation): "
          f"{len(candidates)} docs")
    print("Sample (lowest 10 SSIM):")
    candidates.sort(key=lambda a: a["min_ssim"]
                                  if a["min_ssim"] is not None else 1.0)
    for a in candidates[:10]:
        print(f"  {a['doc_id'][:30]:30s}  kern={a['effective_kern']}  "
              f"cs={a['charSpacingControl']}  ssim={a['min_ssim']}")

    out = {
        "total": total,
        "charSpacingControl_distribution": dict(cs_dist),
        "cross_tab_kern_cs": {f"{k}|{cs}": n for (k, cs), n
                                in cross.items()},
        "audit": audit,
    }
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {RESULT_PATH}")


if __name__ == "__main__":
    main()
