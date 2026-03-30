"""Structural DML diff: compare Word COM layout vs Oxi layout_json.

Usage:
  python dml_diff.py <docx_path>                    # single file diff
  python dml_diff.py <docx_dir> [--summary]          # batch summary

Requires: Word DML cache in pipeline_data/word_dml/ (run word_dml_extract.py first)
"""
import json
import subprocess
import sys
import os
from pathlib import Path
from collections import defaultdict

CACHE_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data", "word_dml")
OXI_ROOT = os.path.join(os.path.dirname(__file__), "..", "..")


def get_oxi_layout(docx_path: str) -> dict:
    """Run Oxi layout_json and parse output."""
    result = subprocess.run(
        ["cargo", "run", "--release", "--example", "layout_json", "--", docx_path],
        capture_output=True, text=True, errors="replace",
        cwd=OXI_ROOT, timeout=120,
    )

    pages = []
    current_page = {"lines": [], "borders": []}
    prev_y = None
    line_chars = 0
    line_x = 0
    line_text = ""

    for raw in result.stdout.split("\n"):
        line = raw.rstrip("\r")
        if line.startswith("PAGE\t"):
            if current_page["lines"]:
                pages.append(current_page)
            parts = line.split("\t")
            current_page = {
                "width": float(parts[2]),
                "height": float(parts[3]),
                "lines": [],
                "borders": [],
            }
            prev_y = None
            line_chars = 0
        elif line.startswith("TEXT\t"):
            parts = line.split("\t")
            y = float(parts[2])
            x = float(parts[1])
            if prev_y is None or abs(y - prev_y) > 0.5:
                if prev_y is not None:
                    current_page["lines"].append({
                        "y": round(prev_y, 2),
                        "x": round(line_x, 2),
                        "chars": line_chars,
                        "text": line_text[:60],
                    })
                line_x = x
                line_chars = 0
                line_text = ""
                prev_y = y
        elif line.startswith("T\t"):
            line_chars += len(line[2:])
            line_text += line[2:]
        elif line.startswith("BORDER\t"):
            parts = line.split("\t")
            current_page["borders"].append({
                "y1": float(parts[2]),
                "y2": float(parts[4]),
                "x1": float(parts[1]),
                "x2": float(parts[3]),
            })

    if prev_y is not None and line_chars > 0:
        current_page["lines"].append({
            "y": round(prev_y, 2),
            "x": round(line_x, 2),
            "chars": line_chars,
            "text": line_text[:60],
        })
    if current_page["lines"]:
        pages.append(current_page)

    return {"pages": pages}


def diff_document(docx_path: str, verbose: bool = True) -> dict:
    """Compare Word DML cache vs Oxi layout for a single document."""
    doc_id = Path(docx_path).stem
    cache_path = os.path.join(CACHE_DIR, f"{doc_id}.json")

    if not os.path.exists(cache_path):
        if verbose:
            print(f"[SKIP] No Word DML cache for {doc_id}")
        return {"status": "no_cache"}

    with open(cache_path, encoding="utf-8") as f:
        word_data = json.load(f)

    if verbose:
        print(f"Running Oxi layout_json...")
    oxi_data = get_oxi_layout(docx_path)

    # Compare pages
    word_pages = word_data["pages"]
    oxi_pages = len(oxi_data["pages"])

    # Flatten Word paragraphs into per-page lines
    word_page_lines = defaultdict(list)
    for p in word_data["paragraphs"]:
        for line in p["lines"]:
            word_page_lines[p["page"]].append(line)

    report = {
        "doc_id": doc_id,
        "word_pages": word_pages,
        "oxi_pages": oxi_pages,
        "page_match": word_pages == oxi_pages,
        "diffs": [],
    }

    # Compare line-by-line per page
    total_y_err = 0
    total_char_diff = 0
    n_compared = 0
    page_diffs = []

    word_page_nums = sorted(word_page_lines.keys())
    for pi, word_pg in enumerate(word_page_nums):
        if pi >= oxi_pages:
            page_diffs.append({
                "page": pi + 1,
                "error": "Oxi missing this page",
            })
            continue

        w_lines = word_page_lines[word_pg]
        o_lines = oxi_data["pages"][pi]["lines"]

        pd = {
            "page": pi + 1,
            "word_lines": len(w_lines),
            "oxi_lines": len(o_lines),
            "line_count_match": len(w_lines) == len(o_lines),
            "y_errors": [],
            "char_diffs": [],
        }

        for li in range(min(len(w_lines), len(o_lines))):
            wl = w_lines[li]
            ol = o_lines[li]
            dy = ol["y"] - wl["y"]
            dc = ol["chars"] - wl["chars"]
            total_y_err += abs(dy)
            total_char_diff += abs(dc)
            n_compared += 1

            if abs(dy) > 0.5 or abs(dc) > 2:
                pd["y_errors"].append({
                    "line": li + 1,
                    "word_y": wl["y"],
                    "oxi_y": ol["y"],
                    "dy": round(dy, 2),
                    "word_chars": wl["chars"],
                    "oxi_chars": ol["chars"],
                    "d_chars": dc,
                })

        page_diffs.append(pd)

    report["page_diffs"] = page_diffs
    if n_compared > 0:
        report["mean_y_error"] = round(total_y_err / n_compared, 3)
        report["mean_char_diff"] = round(total_char_diff / n_compared, 3)
    else:
        report["mean_y_error"] = 999
        report["mean_char_diff"] = 999

    # Print report
    if verbose:
        print(f"\n{'='*60}")
        print(f"DML DIFF: {doc_id}")
        print(f"{'='*60}")
        print(f"Pages: Word={word_pages}, Oxi={oxi_pages} {'OK' if report['page_match'] else 'NG'}")
        if n_compared > 0:
            print(f"Mean |dy|: {report['mean_y_error']:.2f}pt")
            print(f"Mean |Δchars|: {report['mean_char_diff']:.2f}")

        for pd in page_diffs:
            if "error" in pd:
                print(f"\n  Page {pd['page']}: {pd['error']}")
                continue

            match = "OK" if pd["line_count_match"] else "NG"
            print(f"\n  Page {pd['page']}: lines Word={pd['word_lines']} Oxi={pd['oxi_lines']} {match}")

            for err in pd["y_errors"][:10]:
                markers = []
                if abs(err["dy"]) > 1:
                    markers.append(f"dy={err['dy']:+.1f}pt")
                if abs(err["d_chars"]) > 2:
                    markers.append(f"Δch={err['d_chars']:+d}")
                print(f"    line {err['line']:3d}: Word y={err['word_y']:7.1f} ({err['word_chars']:3d}ch) "
                      f"Oxi y={err['oxi_y']:7.1f} ({err['oxi_chars']:3d}ch) {' '.join(markers)}")

            if len(pd["y_errors"]) > 10:
                print(f"    ... and {len(pd['y_errors'])-10} more")

    return report


def batch_summary(docx_dir: str):
    """Run diff on all documents with cached Word DML."""
    results = []
    docx_files = sorted(Path(docx_dir).glob("*.docx"))

    for f in docx_files:
        doc_id = f.stem
        cache_path = os.path.join(CACHE_DIR, f"{doc_id}.json")
        if not os.path.exists(cache_path):
            continue
        try:
            report = diff_document(str(f), verbose=False)
            if report.get("status") == "no_cache":
                continue
            results.append(report)
        except Exception as e:
            print(f"  [ERROR] {doc_id}: {e}")

    # Sort by mean_y_error descending
    results.sort(key=lambda r: -r.get("mean_y_error", 0))

    print(f"\n{'='*60}")
    print(f"DML DIFF SUMMARY ({len(results)} documents)")
    print(f"{'='*60}")
    print(f"{'Document':40s} {'Pages':>8s} {'|dy|':>7s} {'|Δch|':>7s}")
    print(f"{'-'*40} {'-'*8} {'-'*7} {'-'*7}")

    total_dy = 0
    total_dc = 0
    n = 0
    for r in results:
        pg = f"{r['oxi_pages']}/{r['word_pages']}"
        dy = r.get("mean_y_error", 999)
        dc = r.get("mean_char_diff", 999)
        marker = " NG" if not r["page_match"] else ""
        print(f"{r['doc_id'][:40]:40s} {pg:>8s} {dy:7.2f} {dc:7.2f}{marker}")
        total_dy += dy
        total_dc += dc
        n += 1

    if n > 0:
        print(f"\n{'Average':40s} {'':>8s} {total_dy/n:7.2f} {total_dc/n:7.2f}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python dml_diff.py <docx_path_or_dir> [--summary]")
        sys.exit(1)

    target = sys.argv[1]

    if os.path.isdir(target):
        batch_summary(target)
    else:
        diff_document(target)


if __name__ == "__main__":
    main()
