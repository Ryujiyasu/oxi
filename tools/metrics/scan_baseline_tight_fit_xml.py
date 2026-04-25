"""Faster scan: directly parse DOCX XML for tight-fit textboxes (XML only, no cargo)."""
import os
import glob
import zipfile
import re

DOCX_DIR = "tools/golden-test/documents/docx"


def scan_xml_tight(docx_path: str):
    """Find textboxes where height < ~30pt (potentially tight-fit single-line)."""
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read('word/document.xml').decode('utf-8', errors='replace')
    except Exception:
        return []

    # Find all wp:anchor elements with their cy + check if it has spAutoFit + text content
    matches = []
    for m in re.finditer(r'<wp:anchor\b[^>]*>(.*?)</wp:anchor>', xml, re.DOTALL):
        block = m.group(0)
        # Get cy
        cy_m = re.search(r'<a:ext\s+cx="(\d+)"\s+cy="(\d+)"', block) or re.search(r'<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"', block)
        if not cy_m:
            continue
        cy_emu = int(cy_m.group(2))
        cy_pt = cy_emu / 12700
        # Tight-fit candidate: height < 35pt (tight for single line)
        if cy_pt > 35:
            continue
        # Has text content?
        has_text = '<w:t' in block
        if not has_text:
            continue
        # Has spAutoFit (means Word auto-fits, which is the key trigger)
        has_autofit = '<a:spAutoFit' in block
        # Get first <w:t> content
        text_m = re.search(r'<w:t[^>]*>([^<]+)</w:t>', block)
        text = text_m.group(1) if text_m else ''
        matches.append((cy_pt, has_autofit, text[:30]))
    return matches


def main():
    files = sorted(glob.glob(os.path.join(DOCX_DIR, '*.docx')))
    print(f"Scanning {len(files)} docs...\n")

    affected = []
    for f in files:
        name = os.path.splitext(os.path.basename(f))[0]
        results = scan_xml_tight(f)
        if results:
            affected.append((name, results))

    print(f"=== Docs with tight-fit candidate textboxes ({len(affected)}) ===")
    for (name, results) in affected:
        autofit_count = sum(1 for r in results if r[1])
        print(f"\n  {name}: {len(results)} small textboxes, {autofit_count} with spAutoFit")
        for (cy, af, text) in results[:3]:
            print(f"    height={cy:.2f}pt autoFit={af} text={text!r}")


if __name__ == "__main__":
    main()
