"""Scan baseline docs for adjacent w:p with DIFFERENT line=exact values.

Target: docs where para N has line=X exact, para N+1 has line=Y exact,
X != Y. These exhibit the boundary +2pt (or more) drift bug.

Output: list of (doc, N, X, Y) candidates for real-doc verification.
"""
import os, sys, re, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC_DIR = r"tools/golden-test/documents/docx"

def scan_doc(path):
    """Return list of (para_idx, lineA, lineB) boundaries where lineA != lineB."""
    try:
        with zipfile.ZipFile(path) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
    except Exception:
        return []
    # Simple regex to find consecutive w:p blocks with line/lineRule.
    # Extract each top-level w:p block (non-nested in w:tbl is hard but we'll accept all).
    # Find line=X lineRule=exact within each w:p block.
    results = []
    # Find all <w:p> blocks and their spacing
    p_blocks = re.findall(r'<w:p[ >](?:(?!<w:p[ >]).)*?</w:p>', xml, re.DOTALL)
    lines = []
    for i, b in enumerate(p_blocks):
        m = re.search(r'<w:spacing[^/]*w:line="(\d+)"[^/]*w:lineRule="exact"', b)
        if not m:
            m = re.search(r'<w:spacing[^/]*w:lineRule="exact"[^/]*w:line="(\d+)"', b)
        lines.append(int(m.group(1)) if m else None)
    # Find adjacent pairs with different exact values
    for i in range(len(lines) - 1):
        a, b = lines[i], lines[i+1]
        if a is not None and b is not None and a != b:
            results.append((i+1, a, b))  # 1-indexed para_idx
    return results


def main():
    if not os.path.isdir(DOC_DIR):
        print(f"docx dir not found: {DOC_DIR}"); return
    hits = []
    for name in sorted(os.listdir(DOC_DIR)):
        if not name.endswith(".docx"): continue
        path = os.path.join(DOC_DIR, name)
        boundaries = scan_doc(path)
        if boundaries:
            hits.append((name, boundaries))
    print(f"Docs with line=exact boundary pairs: {len(hits)}")
    for name, bounds in hits[:30]:
        b0 = bounds[0]
        print(f"  {name[:50]}: {len(bounds)} boundaries, first at para {b0[0]}: line={b0[1]}→{b0[2]}")


if __name__ == "__main__":
    main()
