"""
Compare Word PDF vs LibreOffice PDF layout for the same docx files.
Shows where LibreOffice diverges from Word's rendering.
"""
import json
import os
import sys

try:
    import fitz
except ImportError:
    print("ERROR: pip install pymupdf")
    sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WORD_PDF_DIR = os.path.join(SCRIPT_DIR, "output", "pdfs")
LIBRE_PDF_DIR = os.path.join(SCRIPT_DIR, "output", "pdfs_libreoffice")
MANIFEST_PATH = os.path.join(SCRIPT_DIR, "docx_tests", "manifest.json")


def extract_main_lines(pdf_path: str) -> list[dict]:
    """Extract main text lines (skip metadata/reference lines)."""
    doc = fitz.open(pdf_path)
    lines = []

    for page in doc:
        rawdict = page.get_text("rawdict")
        for block in rawdict.get("blocks", []):
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                text = ""
                chars_data = []
                font_sizes = set()
                for span in line["spans"]:
                    font_sizes.add(round(span["size"], 2))
                    for ci in span.get("chars", []):
                        text += ci["c"]
                        chars_data.append({
                            "c": ci["c"],
                            "x0": round(ci["bbox"][0], 2),
                            "y0": round(ci["bbox"][1], 2),
                            "x1": round(ci["bbox"][2], 2),
                            "y1": round(ci["bbox"][3], 2),
                            "w": round(ci["bbox"][2] - ci["bbox"][0], 2),
                        })

                # Skip metadata
                if text.startswith("[TEST]") or text.startswith("[REFERENCE"):
                    continue
                if not chars_data:
                    continue
                # Skip single-char reference lines (from generate_test_docx)
                if len(chars_data) <= 2 and len(text.strip()) <= 1:
                    continue

                bbox = line["bbox"]
                lines.append({
                    "text": text,
                    "char_count": len(chars_data),
                    "y0": round(bbox[1], 2),
                    "y1": round(bbox[3], 2),
                    "x0": round(bbox[0], 2),
                    "x1": round(bbox[2], 2),
                    "line_height": round(bbox[3] - bbox[1], 2),
                    "line_width": round(bbox[2] - bbox[0], 2),
                    "font_sizes": sorted(font_sizes),
                    "chars": chars_data,
                })

    doc.close()
    return lines


def compare_layouts(word_lines, libre_lines) -> dict:
    """Compare Word and LibreOffice line-by-line."""
    diffs = {
        "line_count_word": len(word_lines),
        "line_count_libre": len(libre_lines),
        "line_count_match": len(word_lines) == len(libre_lines),
        "line_break_diffs": [],
        "height_diffs": [],
        "width_diffs": [],
        "y_position_diffs": [],
    }

    # Compare line breaks (same text on same line?)
    min_lines = min(len(word_lines), len(libre_lines))
    for i in range(min_lines):
        wl = word_lines[i]
        ll = libre_lines[i]

        if wl["text"] != ll["text"]:
            diffs["line_break_diffs"].append({
                "line": i + 1,
                "word_text": wl["text"][:60],
                "libre_text": ll["text"][:60],
                "word_chars": wl["char_count"],
                "libre_chars": ll["char_count"],
            })

        h_diff = ll["line_height"] - wl["line_height"]
        if abs(h_diff) > 0.1:
            diffs["height_diffs"].append({
                "line": i + 1,
                "word_h": wl["line_height"],
                "libre_h": ll["line_height"],
                "diff": round(h_diff, 2),
            })

        w_diff = ll["line_width"] - wl["line_width"]
        if abs(w_diff) > 0.5:
            diffs["width_diffs"].append({
                "line": i + 1,
                "word_w": wl["line_width"],
                "libre_w": ll["line_width"],
                "diff": round(w_diff, 2),
            })

        y_diff = ll["y0"] - wl["y0"]
        if abs(y_diff) > 0.5:
            diffs["y_position_diffs"].append({
                "line": i + 1,
                "word_y": wl["y0"],
                "libre_y": ll["y0"],
                "diff": round(y_diff, 2),
            })

    return diffs


def main():
    with open(MANIFEST_PATH, encoding="utf-8") as f:
        manifest = json.load(f)

    results = []

    print(f"{'File':<40} {'Lines W/L':>9} {'Break':>5} {'HtDiff':>6} {'YDiff':>6} {'WDiff':>6}")
    print("-" * 80)

    for entry in manifest:
        pdf_name = entry["filename"].replace(".docx", ".pdf")
        word_pdf = os.path.join(WORD_PDF_DIR, pdf_name)
        libre_pdf = os.path.join(LIBRE_PDF_DIR, pdf_name)

        if not os.path.exists(word_pdf) or not os.path.exists(libre_pdf):
            continue

        try:
            word_lines = extract_main_lines(word_pdf)
            libre_lines = extract_main_lines(libre_pdf)
            diffs = compare_layouts(word_lines, libre_lines)
            diffs["test_case"] = entry
            results.append(diffs)

            breaks = len(diffs["line_break_diffs"])
            h_diffs = len(diffs["height_diffs"])
            y_diffs = len(diffs["y_position_diffs"])
            w_diffs = len(diffs["width_diffs"])

            status = ""
            if breaks > 0:
                status += " BREAK!"
            if h_diffs > 0:
                status += " HEIGHT"
            if y_diffs > 0:
                status += " Y-POS"

            print(f"{entry['filename']:<40} {diffs['line_count_word']:>4}/{diffs['line_count_libre']:<4} "
                  f"{breaks:>5} {h_diffs:>6} {y_diffs:>6} {w_diffs:>6}{status}")

        except Exception as e:
            print(f"{entry['filename']:<40} ERROR: {e}")

    # Write detailed results
    out_path = os.path.join(SCRIPT_DIR, "output", "word_vs_libre.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f"\nDetailed results: {out_path}")

    # Show the worst line break diffs
    print("\n=== LINE BREAK DIFFERENCES (where LibreOffice breaks differently) ===")
    for r in results:
        if r["line_break_diffs"]:
            tc = r["test_case"]
            print(f"\n--- {tc['font']} {tc['size_pt']}pt {tc['lang']} ---")
            for d in r["line_break_diffs"][:3]:
                print(f"  Line {d['line']}: Word [{d['word_chars']}ch] \"{d['word_text']}\"")
                print(f"          Libre [{d['libre_chars']}ch] \"{d['libre_text']}\"")

    # Show Y position drifts
    print("\n=== Y-POSITION DRIFT (cumulative layout differences) ===")
    for r in results:
        if r["y_position_diffs"]:
            tc = r["test_case"]
            last = r["y_position_diffs"][-1]
            print(f"  {tc['font']:<16} {tc['size_pt']:>5}pt {tc['lang']:<6} "
                  f"max Y drift: {last['diff']:+.2f}pt at line {last['line']}")


if __name__ == "__main__":
    main()
