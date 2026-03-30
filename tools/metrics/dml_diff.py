"""Structural DML diff: compare Word COM layout vs Oxi layout_json.

Usage:
  python dml_diff.py <docx_path>                    # single file diff
  python dml_diff.py <docx_dir> [--summary]          # batch summary

Requires: Word DML cache in pipeline_data/word_dml/ (run word_dml_extract.py first)
Uses: cargo run --release --example layout_json -- <docx> --structure
"""
import json
import subprocess
import sys
import os
from pathlib import Path
from collections import defaultdict

CACHE_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data", "word_dml")
OXI_ROOT = os.path.join(os.path.dirname(__file__), "..", "..")


def get_oxi_structure(docx_path: str) -> dict:
    """Run Oxi layout_json --structure and parse output."""
    result = subprocess.run(
        ["cargo", "run", "--release", "--example", "layout_json", "--", docx_path, "--structure"],
        capture_output=True, text=True, errors="replace",
        cwd=OXI_ROOT, timeout=120,
    )

    pages = []
    current_page = {"paragraphs": [], "table_rows": []}

    current_para = None  # {"index": N, "y": Y, "lines": [...]}

    for raw in result.stdout.split("\n"):
        line = raw.rstrip("\r")
        if line.startswith("PAGE\t"):
            if current_para:
                current_page["paragraphs"].append(current_para)
                current_para = None
            if pages or current_page["paragraphs"] or current_page["table_rows"]:
                pages.append(current_page)
            parts = line.split("\t")
            current_page = {
                "width": float(parts[2]),
                "height": float(parts[3]),
                "paragraphs": [],
                "table_rows": [],
            }
        elif line.startswith("PARA\t"):
            if current_para:
                current_page["paragraphs"].append(current_para)
            parts = line.split("\t")
            idx = int(parts[1])
            y = float(parts[2].split("=")[1])
            current_para = {"index": idx, "y": y, "lines": []}
        elif line.startswith("  LINE\t"):
            parts = line.strip().split("\t")
            ly = float(parts[1].split("=")[1])
            lc = int(parts[2].split("=")[1])
            if current_para:
                current_para["lines"].append({"y": ly, "chars": lc})
        elif line.startswith("  ROW\t"):
            parts = line.strip().split("\t")
            ri = int(parts[1])
            ry = float(parts[2].split("=")[1])
            rh = float(parts[3].split("=")[1])
            current_page["table_rows"].append({"row": ri, "y": ry, "h": rh})
        elif line.startswith("TABLE_START"):
            if current_para:
                current_page["paragraphs"].append(current_para)
                current_para = None

    if current_para:
        current_page["paragraphs"].append(current_para)
    if current_page["paragraphs"] or current_page["table_rows"]:
        pages.append(current_page)

    return {"pages": pages}


def get_word_structure(cache_path: str) -> dict:
    """Parse Word DML cache into paragraph/table structure."""
    with open(cache_path, encoding="utf-8") as f:
        data = json.load(f)

    pages_dict = defaultdict(lambda: {"paragraphs": [], "table_rows": []})

    # Paragraphs: group lines by paragraph, split by page
    for p in data.get("paragraphs", []):
        pg = p["page"]
        para = {
            "index": p["index"],
            "y": p["y"],
            "lines": p.get("lines", [{"y": p["y"], "chars": len(p.get("text", ""))}]),
        }
        pages_dict[pg]["paragraphs"].append(para)

    # Tables: extract row Y positions per page
    # Build page→max_y map from paragraphs to determine page boundaries
    page_max_y = {}
    for p in data.get("paragraphs", []):
        pg = p["page"]
        page_max_y[pg] = max(page_max_y.get(pg, 0), p["y"])

    for t in data.get("tables", []):
        for rd in t.get("row_data", []):
            row_y = rd["y"]
            # Find page: row belongs to the page where its Y falls within range
            # A row at y < margin_top (typ 72pt) on a later page has small y
            # Use nearest paragraph page as reference
            pg = 1
            best_dist = float("inf")
            for p in data.get("paragraphs", []):
                dist = abs(p["y"] - row_y)
                if dist < best_dist:
                    best_dist = dist
                    pg = p["page"]
            pages_dict[pg]["table_rows"].append({
                "row": len(pages_dict[pg]["table_rows"]),
                "y": row_y,
            })

    pages = [pages_dict[pg] for pg in sorted(pages_dict.keys())]
    return {"pages": pages, "total_pages": data.get("pages", len(pages))}


def diff_document(docx_path: str, verbose: bool = True) -> dict:
    """Compare Word DML cache vs Oxi structure for a single document."""
    doc_id = Path(docx_path).stem
    cache_path = os.path.join(CACHE_DIR, f"{doc_id}.json")

    if not os.path.exists(cache_path):
        if verbose:
            print(f"[SKIP] No Word DML cache for {doc_id}")
        return {"status": "no_cache"}

    word = get_word_structure(cache_path)
    if verbose:
        print(f"Running Oxi layout_json --structure...")
    oxi = get_oxi_structure(docx_path)

    word_pages = word.get("total_pages", len(word["pages"]))
    oxi_pages = len(oxi["pages"])

    report = {
        "doc_id": doc_id,
        "word_pages": word_pages,
        "oxi_pages": oxi_pages,
        "page_match": word_pages == oxi_pages,
        "para_diffs": [],
        "table_row_diffs": [],
        "mean_para_dy": 999,
        "mean_line_dchar": 999,
        "mean_row_dy": 999,
    }

    # === Compare paragraphs (nearest-Y match) ===
    total_para_dy = 0
    total_line_dchar = 0
    n_para = 0
    n_lines = 0

    for pi in range(min(len(word["pages"]), oxi_pages)):
        w_paras = [p for p in word["pages"][pi]["paragraphs"] if p.get("lines")]
        o_paras = oxi["pages"][pi]["paragraphs"]

        # Match Word paragraphs to nearest Oxi paragraph by Y coordinate
        used_oxi = set()
        for wp in w_paras:
            # Find nearest Oxi paragraph not yet used
            best_oi = None
            best_dy = float("inf")
            for oi, op in enumerate(o_paras):
                if oi in used_oxi:
                    continue
                dy = abs(op["y"] - wp["y"])
                if dy < best_dy:
                    best_dy = dy
                    best_oi = oi
            if best_oi is None:
                continue
            used_oxi.add(best_oi)
            op = o_paras[best_oi]
            dy = op["y"] - wp["y"]
            total_para_dy += abs(dy)
            n_para += 1

            # Compare lines within paragraph
            w_lines = wp.get("lines", [])
            o_lines = op.get("lines", [])
            line_diffs = []
            for li in range(min(len(w_lines), len(o_lines))):
                wl = w_lines[li]
                ol = o_lines[li]
                ldy = ol["y"] - wl["y"]
                dch = ol["chars"] - wl["chars"]
                total_line_dchar += abs(dch)
                n_lines += 1
                if abs(dch) > 0 or abs(ldy) > 1.0:
                    line_diffs.append({
                        "line": li + 1,
                        "word_y": wl["y"], "oxi_y": ol["y"], "dy": round(ldy, 2),
                        "word_chars": wl["chars"], "oxi_chars": ol["chars"], "dch": dch,
                    })

            if abs(dy) > 1.0 or len(w_lines) != len(o_lines) or line_diffs:
                report["para_diffs"].append({
                    "page": pi + 1,
                    "para": wp["index"],
                    "dy": round(dy, 2),
                    "word_lines": len(w_lines),
                    "oxi_lines": len(o_lines),
                    "line_diffs": line_diffs,
                })

    if n_para > 0:
        report["mean_para_dy"] = round(total_para_dy / n_para, 2)
    if n_lines > 0:
        report["mean_line_dchar"] = round(total_line_dchar / n_lines, 2)

    # === Compare table rows ===
    total_row_dy = 0
    n_rows = 0
    for pi in range(min(len(word["pages"]), oxi_pages)):
        w_rows = word["pages"][pi]["table_rows"]
        o_rows = oxi["pages"][pi]["table_rows"]
        for ri in range(min(len(w_rows), len(o_rows))):
            wr = w_rows[ri]
            orr = o_rows[ri]
            dy = orr["y"] - wr["y"]
            total_row_dy += abs(dy)
            n_rows += 1
            if abs(dy) > 0.5:
                report["table_row_diffs"].append({
                    "page": pi + 1, "row": ri,
                    "word_y": wr["y"], "oxi_y": orr["y"], "dy": round(dy, 2),
                })

    if n_rows > 0:
        report["mean_row_dy"] = round(total_row_dy / n_rows, 2)

    # Print report
    if verbose:
        print(f"\n{'='*60}")
        print(f"DML DIFF: {doc_id}")
        print(f"{'='*60}")
        pg_status = "OK" if report["page_match"] else "NG"
        print(f"Pages: Word={word_pages}, Oxi={oxi_pages} {pg_status}")
        print(f"Paragraphs: |dy|={report['mean_para_dy']:.2f}pt, |dch|={report['mean_line_dchar']:.2f}")
        if n_rows > 0:
            print(f"Table rows: |dy|={report['mean_row_dy']:.2f}pt ({n_rows} rows)")

        # Show worst paragraph diffs
        for pd in report["para_diffs"][:15]:
            lines_status = "OK" if pd["word_lines"] == pd["oxi_lines"] else f"W={pd['word_lines']} O={pd['oxi_lines']}"
            print(f"  P{pd['para']:3d} dy={pd['dy']:+6.1f}  lines={lines_status}")
            for ld in pd["line_diffs"][:5]:
                print(f"    L{ld['line']}: W y={ld['word_y']:.1f} ({ld['word_chars']}ch) O y={ld['oxi_y']:.1f} ({ld['oxi_chars']}ch) dch={ld['dch']:+d}")

        # Show table row diffs
        for td in report["table_row_diffs"][:10]:
            print(f"  TblRow {td['row']}: W y={td['word_y']:.1f} O y={td['oxi_y']:.1f} dy={td['dy']:+.1f}")

    return report


def batch_summary(docx_dir: str):
    """Run diff on all documents with cached Word DML."""
    results = []
    docx_files = sorted(Path(docx_dir).glob("*.docx"))

    for f in docx_files:
        doc_id = f.stem
        if doc_id.startswith("~$"):
            continue
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

    results.sort(key=lambda r: -(r.get("mean_para_dy", 0) + r.get("mean_row_dy", 0)))

    print(f"\n{'='*70}")
    print(f"DML STRUCTURAL DIFF SUMMARY ({len(results)} documents)")
    print(f"{'='*70}")
    print(f"{'Document':40s} {'Pages':>7s} {'P|dy|':>7s} {'|dch|':>7s} {'R|dy|':>7s}")
    print(f"{'-'*40} {'-'*7} {'-'*7} {'-'*7} {'-'*7}")

    tp = 0; tl = 0; tr = 0; n = 0
    for r in results:
        pg = f"{r['oxi_pages']}/{r['word_pages']}"
        pdy = r.get("mean_para_dy", 999)
        ldch = r.get("mean_line_dchar", 999)
        rdy = r.get("mean_row_dy", 999)
        marker = " NG" if not r["page_match"] else ""
        rdy_str = f"{rdy:7.2f}" if rdy < 999 else "      -"
        print(f"{r['doc_id'][:40]:40s} {pg:>7s} {pdy:7.2f} {ldch:7.2f} {rdy_str}{marker}")
        if pdy < 999: tp += pdy
        if ldch < 999: tl += ldch
        if rdy < 999: tr += rdy
        n += 1

    if n > 0:
        print(f"\n{'Average':40s} {'':>7s} {tp/n:7.2f} {tl/n:7.2f} {tr/n:7.2f}")


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
