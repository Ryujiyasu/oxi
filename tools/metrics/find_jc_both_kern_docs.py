"""Find docs in baseline with jc=both AND effective kern present.

For Mech 2 (justify-time slack) to fire, need both:
  - jc=both alignment somewhere in the doc
  - effective kern (=Mech 1/2 active)
"""
import json
import re
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = Path("tools/golden-test/documents/docx")
KERN = Path("pipeline_data/kern_audit_2026-05-02.json")


def main():
    kern_data = json.loads(KERN.read_text(encoding="utf-8"))
    kern_by_short = {d["doc_id_short"]: d for d in kern_data["audit"]}

    candidates = []
    for f in sorted(DOCX_DIR.glob("*.docx")):
        sid = f.stem[:12]
        kern_info = kern_by_short.get(sid)
        if not kern_info or not kern_info.get("effective_kern"):
            continue
        try:
            with zipfile.ZipFile(f) as z:
                doc_xml = z.read("word/document.xml").decode("utf-8", errors="replace")
        except Exception:
            continue
        # Count jc=both paragraphs
        jc_both_count = len(re.findall(r'<w:pPr>(?:[^<]|<(?!/w:pPr>))*?<w:jc\s+w:val="both"', doc_xml))
        if jc_both_count == 0:
            continue
        # Approx para count
        n_paras = len(re.findall(r"<w:p\b", doc_xml))
        candidates.append({
            "doc_id_short": sid,
            "doc_full": f.stem,
            "kern_val": kern_info.get("effective_kern"),
            "kern_source": "Normal style" if kern_info.get("normal_kern") else "docDefaults",
            "jc_both_count": jc_both_count,
            "n_paras": n_paras,
            "min_ssim": kern_info.get("min_ssim"),
        })

    candidates.sort(key=lambda r: -r["jc_both_count"])
    print(f"\n=== jc=both + kern docs ({len(candidates)} found) ===")
    print(f"{'doc_id':12} kern  jc_both n_paras min_ssim")
    for r in candidates[:20]:
        ssim = f"{r['min_ssim']:.3f}" if r["min_ssim"] else "-"
        print(f"{r['doc_id_short']:12} {r['kern_val']:>4}  {r['jc_both_count']:>6} {r['n_paras']:>6} {ssim}  {r['doc_full']}")


if __name__ == "__main__":
    main()
