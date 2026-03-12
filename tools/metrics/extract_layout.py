"""
Extract layout information from Word-generated PDFs:
- Line break positions (which characters appear on which line)
- Line Y positions (baseline)
- Line heights (gap between baselines)
- Page margins as rendered

This captures Word's actual layout decisions, not just glyph metrics.
"""
import json
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("ERROR: pip install pymupdf")
    sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(SCRIPT_DIR, "output", "pdfs")
OUTPUT_PATH = os.path.join(SCRIPT_DIR, "output", "layout_analysis.json")


def extract_layout(pdf_path: str) -> dict:
    """Extract line-by-line layout from a Word PDF."""
    doc = fitz.open(pdf_path)
    pages = []

    for page_num, page in enumerate(doc):
        rawdict = page.get_text("rawdict")
        lines_data = []

        for block in rawdict.get("blocks", []):
            if block["type"] != 0:  # text block only
                continue

            block_bbox = block["bbox"]  # (x0, y0, x1, y1)

            for line in block["lines"]:
                chars = []
                line_text = ""
                font_sizes = []
                font_names = []

                for span in line["spans"]:
                    font_size = span["size"]
                    font_name = span["font"]

                    span_text = ""
                    for char_info in span.get("chars", []):
                        c = char_info["c"]
                        bbox = char_info["bbox"]
                        chars.append({
                            "char": c,
                            "x0": round(bbox[0], 3),
                            "y0": round(bbox[1], 3),
                            "x1": round(bbox[2], 3),
                            "y1": round(bbox[3], 3),
                            "width": round(bbox[2] - bbox[0], 3),
                        })
                        span_text += c

                    line_text += span_text
                    font_sizes.append(font_size)
                    font_names.append(font_name)

                # Skip metadata lines
                if line_text.startswith("[TEST]") or line_text.startswith("[REFERENCE"):
                    continue

                if not chars:
                    continue

                line_bbox = line["bbox"]
                lines_data.append({
                    "text": line_text,
                    "char_count": len(chars),
                    "line_bbox": {
                        "x0": round(line_bbox[0], 3),
                        "y0": round(line_bbox[1], 3),
                        "x1": round(line_bbox[2], 3),
                        "y1": round(line_bbox[3], 3),
                    },
                    "line_height": round(line_bbox[3] - line_bbox[1], 3),
                    "line_width": round(line_bbox[2] - line_bbox[0], 3),
                    "font_sizes": list(set(font_sizes)),
                    "font_names": list(set(font_names)),
                    "first_char_x": chars[0]["x0"],
                    "last_char_x1": chars[-1]["x0"] + chars[-1]["width"],
                    "chars": chars,
                })

        page_rect = page.rect
        pages.append({
            "page_num": page_num,
            "width": round(page_rect.width, 3),
            "height": round(page_rect.height, 3),
            "line_count": len(lines_data),
            "lines": lines_data,
        })

    doc.close()

    # Compute inter-line spacing
    for page_data in pages:
        lines = page_data["lines"]
        for i in range(1, len(lines)):
            prev = lines[i - 1]
            curr = lines[i]
            curr["baseline_gap"] = round(curr["line_bbox"]["y0"] - prev["line_bbox"]["y0"], 3)

    return {"pages": pages}


def main():
    manifest_path = os.path.join(SCRIPT_DIR, "docx_tests", "manifest.json")
    with open(manifest_path, encoding="utf-8") as f:
        manifest = json.load(f)

    results = []

    for entry in manifest:
        pdf_name = entry["filename"].replace(".docx", ".pdf")
        pdf_path = os.path.join(PDF_DIR, pdf_name)

        if not os.path.exists(pdf_path):
            continue

        print(f"Extracting layout: {pdf_name} ... ", end="", flush=True)

        try:
            layout = extract_layout(pdf_path)
            result = {"test_case": entry, "layout": layout}
            results.append(result)

            # Summary
            total_lines = sum(p["line_count"] for p in layout["pages"])
            main_lines = [l for p in layout["pages"] for l in p["lines"]
                         if not any(c["char"] in "[]=1" for c in l["chars"][:3])]

            if main_lines:
                heights = [l["line_height"] for l in main_lines]
                widths = [l["line_width"] for l in main_lines]
                print(f"{len(main_lines)} lines, "
                      f"height={min(heights):.2f}-{max(heights):.2f}pt, "
                      f"width={min(widths):.1f}-{max(widths):.1f}pt")

                # Show line break points
                for i, l in enumerate(main_lines[:3]):
                    text = l["text"][:40]
                    print(f"    L{i+1}: [{l['char_count']:2d}ch] {text}...")
            else:
                print(f"{total_lines} lines")

        except Exception as e:
            print(f"FAILED: {e}")
            import traceback
            traceback.print_exc()

    # Write results
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f"\nLayout analysis written to {OUTPUT_PATH}")

    # Summary table: line heights by font/size
    print("\n" + "=" * 90)
    print(f"{'Font':<16} {'Size':>5} {'Lang':<6} {'Lines':>5} "
          f"{'LineH':>7} {'BaseGap':>8} {'1stCharX':>8} {'PageW':>6}")
    print("-" * 90)

    for r in results:
        tc = r["test_case"]
        pages = r["layout"]["pages"]
        all_lines = [l for p in pages for l in p["lines"]
                    if not any(c["char"] in "[]=1" for c in l["chars"][:3])]

        if not all_lines:
            continue

        line_heights = [l["line_height"] for l in all_lines]
        baseline_gaps = [l.get("baseline_gap", 0) for l in all_lines if l.get("baseline_gap")]
        first_x = all_lines[0]["first_char_x"] if all_lines else 0
        page_w = pages[0]["width"] if pages else 0

        avg_h = sum(line_heights) / len(line_heights)
        avg_gap = sum(baseline_gaps) / len(baseline_gaps) if baseline_gaps else 0

        print(f"{tc['font']:<16} {tc['size_pt']:>5} {tc['lang']:<6} "
              f"{len(all_lines):>5} {avg_h:>7.2f} {avg_gap:>8.2f} "
              f"{first_x:>8.2f} {page_w:>6.1f}")

    print("=" * 90)


if __name__ == "__main__":
    main()
