"""
Extract glyph metrics from Word-generated PDFs and compare with HarfBuzz.

This script:
1. Reads PDFs produced by word_to_pdf.ps1
2. Extracts glyph positions, widths, and line break points
3. Runs the same text through HarfBuzz with the same font/size
4. Produces a diff table showing Word's adjustments

Requirements:
    pip install pymupdf uharfbuzz fonttools

Usage:
    python extract_metrics.py
    python extract_metrics.py --pdf-dir output/pdfs --font-dir C:/Windows/Fonts
"""

import argparse
import json
import os
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None
    print("WARNING: pymupdf not installed. PDF extraction will be unavailable.")
    print("  Install with: pip install pymupdf")

try:
    import uharfbuzz as hb
except ImportError:
    hb = None
    print("WARNING: uharfbuzz not installed. HarfBuzz comparison will be unavailable.")
    print("  Install with: pip install uharfbuzz")

try:
    from fontTools.ttLib import TTFont
except ImportError:
    TTFont = None
    print("WARNING: fonttools not installed. Font inspection will be unavailable.")
    print("  Install with: pip install fonttools")


# --- Font name to file mapping (Windows) ---
FONT_FILE_MAP = {
    "游明朝": "yumin.ttf",
    "游ゴシック": "yugothic.ttf",  # or YuGothR.ttc
    "Century": "CENTURY.TTF",
    "Times New Roman": "times.ttf",
}


@dataclass
class GlyphInfo:
    """Position and metrics of a single glyph as rendered."""
    char: str
    x: float  # horizontal position (pt)
    y: float  # vertical position (pt)
    width: float  # advance width (pt)
    font_size: float


@dataclass
class LineInfo:
    """A line of text as laid out by the renderer."""
    y: float
    glyphs: list = field(default_factory=list)

    @property
    def text(self):
        return "".join(g.char for g in self.glyphs)


@dataclass
class MetricsDiff:
    """Difference between Word and HarfBuzz for a single glyph."""
    char: str
    word_width: float
    harfbuzz_width: float
    diff: float  # word - harfbuzz
    diff_pct: float  # percentage difference


def extract_glyphs_from_pdf(pdf_path: str) -> list[GlyphInfo]:
    """Extract glyph positions from a PDF using PyMuPDF."""
    if fitz is None:
        raise RuntimeError("pymupdf is required for PDF extraction")

    doc = fitz.open(pdf_path)
    glyphs = []

    for page in doc:
        # Get text with position information
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

        for block in blocks.get("blocks", []):
            if block["type"] != 0:  # text block
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"]
                    font_size = span["size"]
                    origin_x = span["origin"][0]
                    origin_y = span["origin"][1]

                    # Skip metadata lines
                    if text.startswith("[TEST]") or text.startswith("[REFERENCE"):
                        continue

                    # For character-level positions, use get_texttrace or
                    # compute from char widths
                    # PyMuPDF gives us span-level origins; for per-char we
                    # need the rawdict mode
                    for i, char in enumerate(text):
                        if char.strip() == "":
                            continue
                        glyphs.append(GlyphInfo(
                            char=char,
                            x=origin_x,  # approximate; refined below
                            y=origin_y,
                            width=0,  # filled in by detailed extraction
                            font_size=font_size,
                        ))

    doc.close()
    return glyphs


def extract_detailed_glyphs(pdf_path: str) -> list[GlyphInfo]:
    """Extract per-character glyph positions using rawdict mode."""
    if fitz is None:
        raise RuntimeError("pymupdf is required for PDF extraction")

    doc = fitz.open(pdf_path)
    glyphs = []

    for page in doc:
        rawdict = page.get_text("rawdict")

        for block in rawdict.get("blocks", []):
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    font_size = span["size"]

                    # Skip metadata
                    text = "".join(c["c"] for c in span["chars"]) if "chars" in span else span.get("text", "")
                    if text.startswith("[TEST]") or text.startswith("[REFERENCE"):
                        continue

                    for char_info in span.get("chars", []):
                        c = char_info["c"]
                        bbox = char_info["bbox"]  # (x0, y0, x1, y1)
                        origin = char_info.get("origin", (bbox[0], bbox[3]))

                        glyphs.append(GlyphInfo(
                            char=c,
                            x=origin[0],
                            y=origin[1],
                            width=bbox[2] - bbox[0],  # x1 - x0
                            font_size=font_size,
                        ))

    doc.close()
    return glyphs


def get_harfbuzz_widths(text: str, font_path: str, font_size: float) -> dict[str, float]:
    """Get character widths from HarfBuzz for comparison."""
    if hb is None:
        raise RuntimeError("uharfbuzz is required for HarfBuzz comparison")

    blob = hb.Blob.from_file_path(font_path)
    face = hb.Face(blob)
    font = hb.Font(face)

    # Scale to match point size
    upem = face.upem
    scale = font_size / upem

    char_widths = {}

    for char in set(text):
        if char.strip() == "":
            continue

        buf = hb.Buffer()
        buf.add_str(char)
        buf.guess_segment_properties()
        hb.shape(font, buf)

        positions = buf.glyph_positions
        if positions:
            # Convert from font units to points
            char_widths[char] = positions[0].x_advance * scale

    return char_widths


def compare_metrics(
    word_glyphs: list[GlyphInfo],
    harfbuzz_widths: dict[str, float],
) -> list[MetricsDiff]:
    """Compare Word glyph widths against HarfBuzz."""
    # Aggregate Word widths per character
    word_char_widths: dict[str, list[float]] = {}
    for g in word_glyphs:
        if g.width > 0:
            word_char_widths.setdefault(g.char, []).append(g.width)

    diffs = []
    for char, widths in sorted(word_char_widths.items()):
        avg_word_width = sum(widths) / len(widths)
        hb_width = harfbuzz_widths.get(char, 0)

        if hb_width > 0:
            diff = avg_word_width - hb_width
            diff_pct = (diff / hb_width) * 100
        else:
            diff = avg_word_width
            diff_pct = 100.0

        diffs.append(MetricsDiff(
            char=char,
            word_width=round(avg_word_width, 4),
            harfbuzz_width=round(hb_width, 4),
            diff=round(diff, 4),
            diff_pct=round(diff_pct, 2),
        ))

    return diffs


def process_test_case(
    pdf_path: str,
    font_path: str,
    text: str,
    font_size: float,
) -> dict:
    """Process a single test case: extract Word metrics, compare with HarfBuzz."""
    # Extract from Word PDF
    word_glyphs = extract_detailed_glyphs(pdf_path)

    # Get HarfBuzz reference
    hb_widths = get_harfbuzz_widths(text, font_path, font_size)

    # Compare
    diffs = compare_metrics(word_glyphs, hb_widths)

    # Summary statistics
    if diffs:
        abs_diffs = [abs(d.diff) for d in diffs]
        max_diff = max(abs_diffs)
        avg_diff = sum(abs_diffs) / len(abs_diffs)
        pct_diffs = [abs(d.diff_pct) for d in diffs]
        max_pct = max(pct_diffs)
        avg_pct = sum(pct_diffs) / len(pct_diffs)
    else:
        max_diff = avg_diff = max_pct = avg_pct = 0

    return {
        "word_glyph_count": len(word_glyphs),
        "unique_chars_compared": len(diffs),
        "max_diff_pt": round(max_diff, 4),
        "avg_diff_pt": round(avg_diff, 4),
        "max_diff_pct": round(max_pct, 2),
        "avg_diff_pct": round(avg_pct, 2),
        "diffs": [asdict(d) for d in diffs],
    }


def main():
    parser = argparse.ArgumentParser(description="Extract and compare Word vs HarfBuzz metrics")
    parser.add_argument("--pdf-dir", default=os.path.join(os.path.dirname(__file__), "output", "pdfs"))
    parser.add_argument("--font-dir", default="C:\\Windows\\Fonts")
    parser.add_argument("--manifest", default=os.path.join(os.path.dirname(__file__), "docx_tests", "manifest.json"))
    parser.add_argument("--output", default=os.path.join(os.path.dirname(__file__), "output", "metrics_diff.json"))
    args = parser.parse_args()

    # Load manifest
    with open(args.manifest, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    results = []

    for entry in manifest:
        pdf_name = entry["filename"].replace(".docx", ".pdf")
        pdf_path = os.path.join(args.pdf_dir, pdf_name)

        if not os.path.exists(pdf_path):
            print(f"SKIP (no PDF): {pdf_name}")
            continue

        font_file = FONT_FILE_MAP.get(entry["font"])
        if not font_file:
            print(f"SKIP (unknown font): {entry['font']}")
            continue

        font_path = os.path.join(args.font_dir, font_file)
        if not os.path.exists(font_path):
            print(f"SKIP (font not found): {font_path}")
            continue

        print(f"Processing: {entry['filename']} ...")

        try:
            result = process_test_case(
                pdf_path=pdf_path,
                font_path=font_path,
                text=entry["text"],
                font_size=entry["size_pt"],
            )
            result["test_case"] = entry
            results.append(result)

            print(f"  Chars compared: {result['unique_chars_compared']}")
            print(f"  Max diff: {result['max_diff_pt']}pt ({result['max_diff_pct']}%)")
            print(f"  Avg diff: {result['avg_diff_pt']}pt ({result['avg_diff_pct']}%)")

        except Exception as e:
            print(f"  ERROR: {e}")
            results.append({"test_case": entry, "error": str(e)})

    # Write results
    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f"\nResults written to {args.output}")

    # Print summary table
    print("\n" + "=" * 80)
    print(f"{'Font':<16} {'Size':>5} {'Lang':<6} {'MaxDiff':>8} {'AvgDiff':>8} {'Max%':>6} {'Avg%':>6}")
    print("-" * 80)
    for r in results:
        if "error" in r:
            continue
        tc = r["test_case"]
        print(f"{tc['font']:<16} {tc['size_pt']:>5} {tc['lang']:<6} "
              f"{r['max_diff_pt']:>7.3f}pt {r['avg_diff_pt']:>7.3f}pt "
              f"{r['max_diff_pct']:>5.1f}% {r['avg_diff_pct']:>5.1f}%")
    print("=" * 80)


if __name__ == "__main__":
    main()
