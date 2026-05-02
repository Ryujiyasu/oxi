"""3a4f9f static analysis — what causes pages 1-8 to accumulate ~1 page drift?

Looking for high-frequency layout features that Oxi might mishandle:
- exact lineRule paragraphs (221 total)
- atLeast lineRule paragraphs (109 total)
- inline tables (108 total)
- drawings / textboxes
- specific styles

Strategy: parse paragraphs in document order, classify each by style/features,
identify which paragraphs are in pages 1-8 (the cumulative-drift region).

Without exact page boundaries, we can approximate: para 179 is the FIRST shifted
paragraph. So paras 1..178 contain the cumulative drift trigger. Tabulate
features in this range vs the rest.
"""
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = "C:/Users/ryuji/oxi-main/tools/golden-test/documents/docx"


def load_doc():
    for f in os.listdir(DOCX):
        if f.startswith("3a4f9f") and f.endswith(".docx"):
            with zipfile.ZipFile(os.path.join(DOCX, f), "r") as zf:
                return zf.read("word/document.xml").decode("utf-8", errors="replace")
    return None


def main():
    doc = load_doc()
    if not doc:
        print("Doc not found")
        return

    body = re.search(r"<w:body>(.*)</w:body>", doc, re.DOTALL).group(1)

    # Find all paragraphs in body order — but paragraphs inside tables also count
    paras = re.findall(r"<w:p\b[^>]*>.*?</w:p>", body, re.DOTALL)
    print(f"Total paragraphs (incl. cells): {len(paras)}")

    # Classify each
    feats = []
    for i, p in enumerate(paras, 1):
        f = {"i": i}
        m = re.search(r'<w:pStyle w:val="([^"]+)"', p)
        f["style"] = m.group(1) if m else None
        m = re.search(r'<w:spacing[^/]*?w:lineRule="([^"]+)"', p)
        f["lineRule"] = m.group(1) if m else None
        m = re.search(r'<w:spacing[^/]*?w:line="(\d+)"', p)
        f["line_val"] = int(m.group(1)) if m else None
        m = re.search(r'<w:sz w:val="(\d+)"', p)
        f["sz_pt"] = int(m.group(1)) / 2 if m else None
        f["has_drawing"] = "<w:drawing" in p
        f["has_pict"] = "<w:pict>" in p
        f["pageBrk"] = '<w:br w:type="page"' in p
        f["empty"] = not re.search(r'<w:t[ >]', p)
        feats.append(f)

    # Pre-179 vs post-179 comparison (drift starts at para 179)
    pre = feats[:178]
    post = feats[178:]

    def tabulate(group, label):
        from collections import Counter
        styles = Counter(f["style"] for f in group)
        line_rules = Counter(f["lineRule"] for f in group)
        sizes = Counter(f["sz_pt"] for f in group)
        empty_count = sum(1 for f in group if f["empty"])
        drawing_count = sum(1 for f in group if f["has_drawing"])
        pict_count = sum(1 for f in group if f["has_pict"])
        pgbrk = sum(1 for f in group if f["pageBrk"])
        print(f"\n=== {label} (n={len(group)}) ===")
        print(f"  Empty paras: {empty_count}")
        print(f"  Drawings: {drawing_count}, Pict: {pict_count}")
        print(f"  Page breaks: {pgbrk}")
        print(f"  Top 10 styles:")
        for s, n in styles.most_common(10):
            print(f"    {s!r}: {n}")
        print(f"  lineRule distribution:")
        for r, n in line_rules.most_common():
            print(f"    {r!r}: {n}")
        print(f"  sz_pt distribution:")
        for s, n in sizes.most_common(8):
            print(f"    {s}: {n}")
        # Density features per paragraph
        density_drawing = drawing_count / max(1, len(group))
        density_empty = empty_count / max(1, len(group))
        print(f"  Drawing density: {density_drawing:.3f} per para")
        print(f"  Empty density: {density_empty:.3f} per para")

    tabulate(pre, "Pre-179 (drift accumulating)")
    tabulate(post[:178], "Post-179 (sample, comparable size)")


if __name__ == "__main__":
    main()
