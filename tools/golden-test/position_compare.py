"""Compare text positions between Oxi layout engine and Word PDF output.

Strategy: Group spans into lines by Y coordinate, then compare line-level
positions (first-char X, line Y) between Oxi and Word.

Uses:
  - `cargo run --release --example layout_json` for Oxi positions
  - PyMuPDF (fitz) for Word PDF text positions
"""
import subprocess
import sys
import json
import fitz  # PyMuPDF
from pathlib import Path
from dataclasses import dataclass, field

SCRIPT_DIR = Path(__file__).parent
PROJECT_ROOT = SCRIPT_DIR.parent.parent
DOCX_DIR = SCRIPT_DIR / "documents" / "docx"
WORD_PDF_DIR = SCRIPT_DIR / "pixel_output" / "word_pdf"
LAYOUT_JSON_BIN = PROJECT_ROOT / "target" / "release" / "examples" / "layout_json.exe"

@dataclass
class TextLine:
    y: float          # top of line
    x_start: float    # leftmost x
    height: float
    text: str         # concatenated text

def get_oxi_lines(docx_path: Path) -> list[TextLine]:
    """Run layout_json and group output into lines by Y coordinate."""
    result = subprocess.run(
        [str(LAYOUT_JSON_BIN), str(docx_path)],
        capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=30
    )
    if result.returncode != 0:
        return []

    # Parse all text spans
    spans = []  # (x, y, height, text)
    pending = None
    for raw_line in result.stdout.splitlines():
        parts = raw_line.split('\t')
        if parts[0] == 'TEXT' and len(parts) >= 14:
            pending = (float(parts[1]), float(parts[2]), float(parts[4]))
        elif parts[0] == 'T' and pending is not None:
            text = parts[1] if len(parts) > 1 else ""
            if text.strip():
                spans.append((*pending, text))
            pending = None
        else:
            pending = None

    # Group into lines: spans within 1pt Y are same line
    spans.sort(key=lambda s: (s[1], s[0]))
    lines = []
    for x, y, h, text in spans:
        if lines and abs(y - lines[-1].y) < 1.0:
            lines[-1].text += text
            lines[-1].x_start = min(lines[-1].x_start, x)
        else:
            lines.append(TextLine(y=y, x_start=x, height=h, text=text))
    return lines

def get_word_lines(pdf_path: Path) -> list[TextLine]:
    """Extract text lines from Word PDF using PyMuPDF."""
    doc = fitz.open(str(pdf_path))
    if len(doc) == 0:
        return []
    page = doc[0]
    blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]

    spans = []  # (x, y, h, text)
    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if not text:
                    continue
                bbox = span["bbox"]
                spans.append((bbox[0], bbox[1], bbox[3] - bbox[1], text))
    doc.close()

    # Group into lines
    spans.sort(key=lambda s: (s[1], s[0]))
    lines = []
    for x, y, h, text in spans:
        if lines and abs(y - lines[-1].y) < 1.0:
            lines[-1].text += text
            lines[-1].x_start = min(lines[-1].x_start, x)
        else:
            lines.append(TextLine(y=y, x_start=x, height=h, text=text))
    return lines

def match_lines(oxi_lines: list[TextLine], word_lines: list[TextLine]):
    """Match lines by text content (prefix match, order-preserving)."""
    matches = []
    word_idx = 0
    for oxi in oxi_lines:
        oxi_text = oxi.text.strip()[:20]
        if not oxi_text or len(oxi_text) < 2:
            continue
        for wi in range(word_idx, min(len(word_lines), word_idx + 10)):
            word_text = word_lines[wi].text.strip()
            # Match if first chars overlap
            if oxi_text[:5] == word_text[:5] or oxi_text in word_text or word_text[:10] in oxi_text:
                matches.append((oxi, word_lines[wi]))
                word_idx = wi + 1
                break
    return matches

def compare_document(docx_path: Path, pdf_path: Path) -> dict | None:
    """Compare line positions for one document."""
    oxi_lines = get_oxi_lines(docx_path)
    word_lines = get_word_lines(pdf_path)

    if not oxi_lines or not word_lines:
        return None

    matches = match_lines(oxi_lines, word_lines)
    if not matches:
        return None

    x_errors = []
    y_errors = []
    details = []

    for oxi, word in matches:
        dx = oxi.x_start - word.x_start
        dy = oxi.y - word.y
        x_errors.append(abs(dx))
        y_errors.append(abs(dy))
        details.append({
            "oxi_text": oxi.text[:40],
            "word_text": word.text[:40],
            "oxi_x": round(oxi.x_start, 2), "oxi_y": round(oxi.y, 2),
            "word_x": round(word.x_start, 2), "word_y": round(word.y, 2),
            "dx": round(dx, 2), "dy": round(dy, 2),
        })

    avg_x = sum(x_errors) / len(x_errors)
    avg_y = sum(y_errors) / len(y_errors)
    max_x = max(x_errors)
    max_y = max(y_errors)

    return {
        "matched": len(matches),
        "oxi_lines": len(oxi_lines),
        "word_lines": len(word_lines),
        "avg_x_err": round(avg_x, 3),
        "avg_y_err": round(avg_y, 3),
        "max_x_err": round(max_x, 3),
        "max_y_err": round(max_y, 3),
        "details": details,
    }

def main():
    if not LAYOUT_JSON_BIN.exists():
        print("Building layout_json example...")
        subprocess.run(
            ["cargo", "build", "--release", "--example", "layout_json"],
            cwd=str(PROJECT_ROOT), check=True
        )

    pairs = []
    for docx in sorted(DOCX_DIR.glob("*.docx")):
        pdf = WORD_PDF_DIR / (docx.stem + ".pdf")
        if pdf.exists():
            pairs.append((docx, pdf))

    if not pairs:
        print("No matching docx/pdf pairs found.")
        sys.exit(1)

    # Filter by argument
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if args:
        target = args[0]
        pairs = [(d, p) for d, p in pairs if target in d.stem]
        if not pairs:
            print(f"No document matching '{target}'")
            sys.exit(1)

    print(f"{'Document':<55} {'Lines':>5} {'AvgX':>7} {'AvgY':>7} {'MaxX':>7} {'MaxY':>7}")
    print("-" * 95)

    all_results = {}
    total_avg_x = []
    total_avg_y = []

    for docx, pdf in pairs:
        name = docx.stem[:52]
        stats = compare_document(docx, pdf)
        if stats is None:
            print(f"  {name:<53} {'SKIP':>5}")
            continue

        print(f"  {name:<53} {stats['matched']:>5} {stats['avg_x_err']:>7.2f} {stats['avg_y_err']:>7.2f} {stats['max_x_err']:>7.2f} {stats['max_y_err']:>7.2f}")
        total_avg_x.append(stats['avg_x_err'])
        total_avg_y.append(stats['avg_y_err'])
        all_results[docx.stem] = stats

    if total_avg_x:
        print("-" * 95)
        overall_x = sum(total_avg_x) / len(total_avg_x)
        overall_y = sum(total_avg_y) / len(total_avg_y)
        print(f"  {'OVERALL':.<53} {'':>5} {overall_x:>7.2f} {overall_y:>7.2f}")

    # Verbose: show per-line details for worst documents
    if "--details" in sys.argv:
        print("\n=== Per-line details (worst Y errors) ===")
        for name, stats in sorted(all_results.items(), key=lambda x: -x[1]['avg_y_err'])[:5]:
            print(f"\n{name} (avg_y={stats['avg_y_err']:.2f}pt, {stats['matched']} lines matched)")
            for d in sorted(stats['details'], key=lambda d: -abs(d['dy']))[:15]:
                print(f"  dy={d['dy']:+7.2f} dx={d['dx']:+7.2f}  oxi=({d['oxi_x']},{d['oxi_y']})  word=({d['word_x']},{d['word_y']})  {d['oxi_text'][:30]}")

    # Save report
    report_path = SCRIPT_DIR / "pixel_output" / "position_comparison.json"
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print(f"\nReport: {report_path}")

if __name__ == "__main__":
    main()
