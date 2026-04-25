"""Scan baseline for docs with tight-fit single-line textboxes.

Tight-fit case: textbox height ≈ inset_t + line_height + inset_b
where the OLD filter would drop the text (line slot extends past clip_bottom).

For each doc, check via Oxi's dump_docx output:
- For each TextBox, check if (height - insets) < line_height + 0.5

Report doc list with affected textbox count.
"""
import os
import glob
import subprocess
import re

DOCX_DIR = "tools/golden-test/documents/docx"


def scan_doc(docx_path: str):
    """Run dump_docx and find tight-fit textboxes."""
    try:
        result = subprocess.run(
            ["cargo", "run", "--release", "--quiet", "--example", "dump_docx", "--", docx_path],
            capture_output=True, text=True, timeout=120, encoding='utf-8', errors='replace'
        )
    except subprocess.TimeoutExpired:
        return None
    out = result.stdout

    # Parse TextBox lines + their content paragraph font_size for line_height estimate
    # TextBox[N]: WIDTHxHEIGHT anchor=N fill=... cr=... blocks=N pos=...
    # Followed by TB[N].P0 ls=Some(LH)/Some("rule") ... | XXpt "text"
    tightfit = 0
    tb_count = 0
    lines = out.split('\n')
    for i, line in enumerate(lines):
        m = re.search(r'TextBox\[(\d+)\]: ([\d.]+)x([\d.]+)', line)
        if m:
            tb_count += 1
            h = float(m.group(3))
            # Find next TB[N].P0 line for line_spacing
            for j in range(i+1, min(i+5, len(lines))):
                m2 = re.search(r'ls=Some\(([\d.]+)\)/(?:Some\("(\w+)"\))', lines[j])
                if m2:
                    line_h = float(m2.group(1))
                    inner_h = h - 7.2  # default insets
                    # Tight-fit: inner_h fits 0 lines but text exists
                    if inner_h < line_h + 0.5 and inner_h > 0:
                        tightfit += 1
                    break
    return (tb_count, tightfit)


def main():
    files = sorted(glob.glob(os.path.join(DOCX_DIR, '*.docx')))
    print(f"Scanning {len(files)} docs (this takes time, ~3s per doc)...\n")

    affected = []
    for i, f in enumerate(files):
        name = os.path.splitext(os.path.basename(f))[0]
        r = scan_doc(f)
        if r is None:
            continue
        (tb, tf) = r
        if tf > 0:
            affected.append((name, tb, tf))
        if (i+1) % 20 == 0:
            print(f"  scanned {i+1}/{len(files)}, found {len(affected)} affected so far")

    print(f"\n=== Docs with tight-fit textboxes ({len(affected)}) ===")
    for (name, tb, tf) in affected:
        print(f"  {name}: {tf}/{tb} textboxes tight-fit")


if __name__ == "__main__":
    main()
