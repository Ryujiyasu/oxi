"""Identify what para_idx=83 in e3c545 is and why Oxi over-reserves it.

Steps:
  1. Read e3c545's document.xml, extract the 84th <w:p> (1-based = idx 83 0-based)
  2. Show its raw XML
  3. From Oxi cached layout: extract all elements with para_idx=83
     to see y_start, y_end, width, n_lines
  4. Check what's neighboring (paras 80-90 to see context)
"""
import json
import re
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = Path("tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx")
LAYOUT = Path("pipeline_data/_e3c545_layout.json")


def extract_paragraphs(doc_xml: str) -> list[str]:
    """Extract <w:p>...</w:p> chunks (top-level body)."""
    # Find all <w:p>...</w:p> at any nesting (regex doesn't track)
    return re.findall(r"<w:p\b[^>]*>(?:.*?)</w:p>", doc_xml, re.DOTALL)


def main():
    with zipfile.ZipFile(DOCX) as z:
        doc_xml = z.read("word/document.xml").decode("utf-8")

    paras = extract_paragraphs(doc_xml)
    print(f"Total <w:p> in document.xml: {len(paras)}")

    # Show paras 80-90 (focused around the singleton page)
    for i in range(80, min(95, len(paras))):
        p = paras[i]
        # Extract pPr style highlights
        pstyle = re.search(r'<w:pStyle\s+w:val="([^"]+)"', p)
        # Extract first text snippet
        texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p)
        full_text = "".join(texts)[:80]
        # Check for special elements
        has_tab = "<w:tab" in p
        has_br = "<w:br" in p
        has_drawing = "<w:drawing" in p
        # Length and run count
        n_runs = len(re.findall(r"<w:r\b", p))
        n_t = len(re.findall(r"<w:t\b", p))
        print(f"\n--- para idx={i} (Word p{i+1}) — pStyle={pstyle.group(1) if pstyle else 'Normal'}, n_runs={n_runs}, n_t={n_t}, len={len(p)} ---")
        if has_tab: print("  has_tab")
        if has_br: print("  has_br")
        if has_drawing: print("  has_drawing")
        print(f"  text: {full_text!r}")
        # Show first 500 chars of XML for the singleton para
        if i == 83:
            print(f"  XML (first 1500 chars):")
            print("  " + p[:1500].replace("\n", "\n  "))

    # Now load Oxi layout and show all elements per para_idx 80-90
    layout = json.loads(LAYOUT.read_text(encoding="utf-8"))
    print(f"\n=== Oxi layout per para_idx (80-90) ===")
    for target_idx in range(80, 95):
        elements = []
        for page in layout["pages"]:
            for el in page["elements"]:
                if el.get("para_idx") == target_idx:
                    elements.append((page["page"], el))
        if not elements:
            print(f"  idx={target_idx}: NO oxi elements")
            continue
        ys = sorted(set(round(el.get("y", 0), 1) for _, el in elements))
        pages = sorted(set(p for p, _ in elements))
        n = len(elements)
        text_sample = "".join(el.get("text", "") for _, el in elements[:30])[:80]
        print(f"  idx={target_idx}: pages={pages}, ys={ys[:5]}{'...' if len(ys)>5 else ''} (n={len(ys)}), n_elems={n}")
        print(f"    text: {text_sample!r}")


if __name__ == "__main__":
    main()
