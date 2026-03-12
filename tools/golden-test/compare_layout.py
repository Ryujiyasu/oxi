"""
Compare Word COM layout positions vs Oxi layout engine output.
Extracts first TEXT element of each "paragraph group" from Oxi output,
then aligns with Word COM TSV output by text matching.
"""

import subprocess
import sys
import os
import re

# Force UTF-8 output
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def parse_oxi_output(lines):
    """Parse Oxi layout_json output, grouping TEXT elements by Y position."""
    paragraphs = []
    current_y = None
    current_texts = []
    current_x = None
    current_height = None
    current_font = None
    current_font_size = None
    page = 0

    i = 0
    while i < len(lines):
        line = lines[i]
        if line.startswith("PAGE\t"):
            parts = line.split("\t")
            page = int(parts[1])
            i += 1
            continue

        if line.startswith("TEXT\t"):
            parts = line.split("\t")
            x = float(parts[1])
            y = float(parts[2])
            w = float(parts[3])
            h = float(parts[4])
            fs = int(parts[5])
            ff = parts[6]

            # Next line is the text content
            if i + 1 < len(lines) and lines[i + 1].startswith("T\t"):
                text = lines[i + 1][2:]
                i += 2
            else:
                text = ""
                i += 1

            if current_y is None or abs(y - current_y) > 0.5:
                # New paragraph line
                if current_y is not None and current_texts:
                    paragraphs.append({
                        "page": page,
                        "y": current_y,
                        "x": current_x,
                        "height": current_height,
                        "font": current_font,
                        "font_size": current_font_size,
                        "text": "".join(current_texts)[:50],
                    })
                current_y = y
                current_x = x
                current_height = h
                current_font = ff
                current_font_size = fs
                current_texts = [text]
            else:
                current_texts.append(text)
        else:
            i += 1
            continue

    # Last paragraph
    if current_y is not None and current_texts:
        paragraphs.append({
            "page": page,
            "y": current_y,
            "x": current_x,
            "height": current_height,
            "font": current_font,
            "font_size": current_font_size,
            "text": "".join(current_texts)[:50],
        })

    return paragraphs


def parse_word_output(lines):
    """Parse Word COM TSV output."""
    paragraphs = []
    for line in lines:
        if line.startswith("idx\t"):
            continue  # header
        parts = line.split("\t")
        if len(parts) < 11:
            continue
        paragraphs.append({
            "idx": int(parts[0]),
            "page": int(parts[1]),
            "y": float(parts[2]),
            "x": float(parts[3]),
            "font": parts[4],
            "font_size": float(parts[5]),
            "space_before": float(parts[6]),
            "space_after": float(parts[7]),
            "line_spacing": float(parts[8]),
            "line_rule": parts[9],
            "text": parts[10] if len(parts) > 10 else "",
        })
    return paragraphs


def main():
    base_dir = r"C:\Users\ryuji\oxi-1"

    files = [
        "tests/fixtures/basic_test.docx",
        "tests/fixtures/comprehensive_test.docx",
    ]

    for docx_file in files:
        full_path = os.path.join(base_dir, docx_file)
        print(f"\n{'='*80}")
        print(f"FILE: {docx_file}")
        print(f"{'='*80}")

        # Run Word COM extraction
        word_result = subprocess.run(
            ["python", os.path.join(base_dir, "tools/golden-test/word_layout_extract.py"), full_path],
            capture_output=True, text=True, cwd=base_dir, encoding="utf-8", errors="replace"
        )
        word_lines = word_result.stdout.strip().split("\n")
        word_paras = parse_word_output(word_lines)

        # Run Oxi layout
        oxi_result = subprocess.run(
            ["cargo", "run", "--example", "layout_json", "--", docx_file],
            capture_output=True, text=True, cwd=base_dir, encoding="utf-8", errors="replace"
        )
        oxi_lines = oxi_result.stdout.strip().split("\n")
        oxi_paras = parse_oxi_output(oxi_lines)

        # Print side-by-side comparison
        print(f"\n{'Word':>4} {'Oxi':>4}  {'W-y':>8} {'O-y':>8} {'diff':>7}  {'W-ht':>5} {'O-ht':>6}  {'W-font':>12} {'O-font':>12}  {'W-sz':>4} {'O-sz':>4}  Text preview")
        print("-" * 120)

        # Match by text content (fuzzy)
        oxi_idx = 0
        for wi, wp in enumerate(word_paras):
            w_text_clean = wp["text"].replace("\\r", "").replace("\\n", "").replace("\\t", "").strip()

            best_match = None
            best_score = 0
            search_range = range(max(0, oxi_idx - 2), min(len(oxi_paras), oxi_idx + 5))

            for oi in search_range:
                op = oxi_paras[oi]
                o_text = op["text"].strip()
                # Compare first 15 chars
                w_prefix = w_text_clean[:15]
                o_prefix = o_text[:15]
                if w_prefix and o_prefix:
                    # Simple character overlap score
                    common = sum(1 for a, b in zip(w_prefix, o_prefix) if a == b)
                    score = common / max(len(w_prefix), len(o_prefix), 1)
                    if score > best_score:
                        best_score = score
                        best_match = oi

            if best_match is not None and best_score > 0.3:
                op = oxi_paras[best_match]
                diff = op["y"] - wp["y"]
                print(f"{wi:4d} {best_match:4d}  {wp['y']:8.2f} {op['y']:8.2f} {diff:+7.2f}  "
                      f"{wp.get('line_spacing', 0):5.1f} {op['height']:6.2f}  "
                      f"{wp['font'][:12]:>12} {op['font'][:12]:>12}  "
                      f"{wp['font_size']:4.0f} {op['font_size']:4d}  "
                      f"{w_text_clean[:40]}")
                oxi_idx = best_match + 1
            else:
                print(f"{wi:4d}    -  {wp['y']:8.2f}     -        -  "
                      f"{wp.get('line_spacing', 0):5.1f}      -  "
                      f"{wp['font'][:12]:>12}            -  "
                      f"{wp['font_size']:4.0f}    -  "
                      f"{w_text_clean[:40]}")

        # Summary statistics
        print(f"\nWord paragraphs: {len(word_paras)}")
        print(f"Oxi text groups: {len(oxi_paras)}")

        # Compute Y diffs for matched pairs
        diffs = []
        oxi_idx = 0
        for wi, wp in enumerate(word_paras):
            w_text_clean = wp["text"].replace("\\r", "").replace("\\n", "").replace("\\t", "").strip()
            best_match = None
            best_score = 0
            search_range = range(max(0, oxi_idx - 2), min(len(oxi_paras), oxi_idx + 5))
            for oi in search_range:
                op = oxi_paras[oi]
                o_text = op["text"].strip()
                w_prefix = w_text_clean[:15]
                o_prefix = o_text[:15]
                if w_prefix and o_prefix:
                    common = sum(1 for a, b in zip(w_prefix, o_prefix) if a == b)
                    score = common / max(len(w_prefix), len(o_prefix), 1)
                    if score > best_score:
                        best_score = score
                        best_match = oi
            if best_match is not None and best_score > 0.3:
                op = oxi_paras[best_match]
                diffs.append(op["y"] - wp["y"])
                oxi_idx = best_match + 1

        if diffs:
            print(f"\nY position differences (Oxi - Word):")
            print(f"  Mean: {sum(diffs)/len(diffs):+.2f} pt")
            print(f"  Max:  {max(diffs):+.2f} pt")
            print(f"  Min:  {min(diffs):+.2f} pt")
            print(f"  Matched: {len(diffs)} pairs")


if __name__ == "__main__":
    main()
