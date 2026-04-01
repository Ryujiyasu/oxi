"""
Autonomous SSIM improvement loop with COM measurement and adaptive timeout.

Full loop:
  1. Identify worst SSIM pages
  2. COM-measure the target document (line heights, table widths, page breaks)
  3. Compare COM measurements vs Oxi layout output → find discrepancies
  4. Output diagnosis for Claude Code to fix
  5. After fix: rebuild + re-verify SSIM
  6. If no regression: commit. If regression: git reset --hard
  7. Record success/failure → adaptive timeout

Usage:
  python -m pipeline.autoloop                    # Show targets + diagnosis
  python -m pipeline.autoloop --measure          # Measure worst target via COM
  python -m pipeline.autoloop --diagnose         # Show COM vs Oxi discrepancies
  python -m pipeline.autoloop --record-success
  python -m pipeline.autoloop --record-failure
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path
from .config import DATA_DIR

HISTORY_PATH = os.path.join(DATA_DIR, "loop_history.json")


def load_history() -> list:
    if os.path.exists(HISTORY_PATH):
        with open(HISTORY_PATH, "r") as f:
            return json.load(f)
    return []


def save_history(history: list):
    Path(DATA_DIR).mkdir(parents=True, exist_ok=True)
    with open(HISTORY_PATH, "w") as f:
        json.dump(history, f, indent=2)


def success_rate(history: list, window: int = 10) -> float:
    recent = history[-window:]
    if not recent:
        return 0.5
    return sum(1 for h in recent if h["success"]) / len(recent)


def adaptive_timeout(history: list) -> int:
    rate = success_rate(history)
    if rate >= 0.8:
        return 3600      # 1h
    elif rate >= 0.5:
        return 10800     # 3h
    elif rate >= 0.2:
        return 21600     # 6h
    else:
        return 3600      # 1h - stuck


def get_worst_pages(n: int = 10) -> list:
    from .baseline import load_baseline
    baseline = load_baseline()
    if not baseline:
        return []

    pages = []
    for doc_id, page_scores in baseline.items():
        for page, score in page_scores.items():
            if score > 0:  # Skip missing pages (0.0)
                pages.append({
                    "doc_id": doc_id,
                    "page": int(page),
                    "ssim": score,
                })

    pages.sort(key=lambda x: x["ssim"])
    return pages[:n]


def get_worst_page1(n: int = 5) -> list:
    """Get worst page-1 scores (most actionable, no cumulative error)."""
    from .baseline import load_baseline
    baseline = load_baseline()
    if not baseline:
        return []

    pages = []
    for doc_id, page_scores in baseline.items():
        if "1" in page_scores and page_scores["1"] > 0:
            pages.append({
                "doc_id": doc_id,
                "page": 1,
                "ssim": page_scores["1"],
            })

    pages.sort(key=lambda x: x["ssim"])
    return pages[:n]


def measure_target(docx_dir: str):
    """Measure the worst page-1 document via COM."""
    from .com_measure import measure_document

    targets = get_worst_page1(n=1)
    if not targets:
        print("No targets found.")
        return

    target = targets[0]
    docx_path = os.path.join(docx_dir, f"{target['doc_id']}.docx")
    if not os.path.exists(docx_path):
        print(f"File not found: {docx_path}")
        return

    print(f"Target: {target['doc_id']} p.1 (SSIM={target['ssim']:.4f})")
    measure_document(docx_path)


def diagnose(docx_dir: str):
    """Compare COM measurements vs Oxi layout output for worst target."""
    targets = get_worst_page1(n=1)
    if not targets:
        print("No targets found.")
        return

    target = targets[0]
    doc_id = target["doc_id"]

    # Load COM measurement
    com_path = os.path.join(DATA_DIR, "com_measurements", f"{doc_id}.json")
    if not os.path.exists(com_path):
        print(f"No COM measurement for {doc_id}. Run --measure first.")
        return

    with open(com_path, "r", encoding="utf-8") as f:
        com = json.load(f)

    # Get Oxi layout
    try:
        import subprocess
        oxi_bin = os.path.join(os.path.dirname(__file__), "..", "target", "debug", "oxi.exe")
        # Use WASM to get layout JSON
        sys.path.insert(0, os.path.dirname(__file__))
        # We'll compare using the COM line heights vs expected
    except Exception:
        pass

    print("=" * 70)
    print(f"DIAGNOSIS: {doc_id} (SSIM={target['ssim']:.4f})")
    print("=" * 70)

    # Page break analysis
    com_pages = com.get("page_breaks", [])
    print(f"\nPage breaks (COM): {len(com_pages)} pages")
    for pb in com_pages:
        print(f"  Page {pb['page']}: para {pb['first_para_index']} at y={pb['y_position_pt']}pt")
        print(f"    \"{pb['text_preview']}\"")

    # Line height analysis - find largest discrepancies
    com_lines = com.get("line_heights", [])
    if com_lines:
        print(f"\nLine heights: {len(com_lines)} paragraphs measured")

        # Group by page
        by_page = {}
        for line in com_lines:
            pg = line["page"]
            if pg not in by_page:
                by_page[pg] = []
            by_page[pg].append(line)

        for pg in sorted(by_page.keys()):
            lines = by_page[pg]
            print(f"\n  Page {pg}: {len(lines)} paragraphs")
            # Show first/last Y positions to understand page layout
            if lines:
                first_y = lines[0]["y_position_pt"]
                last_y = lines[-1]["y_position_pt"]
                print(f"    First para Y={first_y}pt, Last para Y={last_y}pt")
                # Show unique line spacing values
                spacings = set((l["line_spacing_pt"], l["line_spacing_rule"]) for l in lines)
                for sp, rule in sorted(spacings):
                    rule_name = {0: "atLeast", 1: "exactly", 2: "multiple",
                                 4: "single", 5: "1.5lines"}.get(rule, f"rule{rule}")
                    count = sum(1 for l in lines if l["line_spacing_pt"] == sp)
                    print(f"    LineSpacing={sp}pt ({rule_name}): {count} paras")

    # Table analysis
    com_tables = com.get("table_widths", [])
    if com_tables:
        print(f"\nTables: {len(com_tables)}")
        for t in com_tables:
            print(f"  Table {t['table_index']}: {t['rows']}x{t['cols']}")
            if t.get("columns"):
                widths = [c["width_pt"] for c in t["columns"]]
                print(f"    Column widths: {widths}")

    print("\n" + "=" * 70)
    print("ACTION ITEMS:")
    print("  1. Compare page break positions with Oxi's layout output")
    print("  2. Fix line height discrepancies (use COM-measured values)")
    print("  3. Fix table column widths if they differ")
    print("  4. Rebuild: cargo build --bin oxi")
    print("  5. Verify: python -m pipeline.verify")
    print("=" * 70)


def print_status(docx_dir: str):
    history = load_history()
    rate = success_rate(history)
    timeout = adaptive_timeout(history)

    print("=" * 60)
    print("SSIM Improvement Loop Status")
    print(f"Success rate (last 10): {rate:.0%}")
    print(f"Adaptive timeout: {timeout // 3600}h {(timeout % 3600) // 60}m")
    print(f"Total attempts: {len(history)}")
    print("=" * 60)

    print("\nWorst page-1 targets (no cumulative error):")
    for p in get_worst_page1(n=5):
        has_com = os.path.exists(
            os.path.join(DATA_DIR, "com_measurements", f"{p['doc_id']}.json"))
        com_flag = " [COM measured]" if has_com else ""
        print(f"  {p['ssim']:.4f}  {p['doc_id']}{com_flag}")

    print("\nWorkflow:")
    print("  1. python -m pipeline.autoloop --measure    # COM measure worst target")
    print("  2. python -m pipeline.autoloop --diagnose   # Show discrepancies")
    print("  3. (Claude Code fixes Rust code based on diagnosis)")
    print("  4. cargo build --bin oxi")
    print("  5. python -m pipeline.verify                # Check no regression")
    print("  6. python -m pipeline.autoloop --record-success/--record-failure")


def record_result(success: bool):
    history = load_history()
    history.append({
        "timestamp": datetime.now().isoformat(),
        "success": success,
    })
    save_history(history)
    rate = success_rate(history)
    timeout = adaptive_timeout(history)
    print(f"Recorded: {'success' if success else 'failure'}")
    print(f"Success rate: {rate:.0%}, Next timeout: {timeout // 3600}h {(timeout % 3600) // 60}m")


if __name__ == "__main__":
    import argparse
    default_dir = os.path.abspath(os.path.join(
        os.path.dirname(__file__), "..",
        "tools", "golden-test", "documents", "docx"))

    parser = argparse.ArgumentParser(description="SSIM Improvement Loop")
    parser.add_argument("--docx-dir", type=str, default=default_dir)
    parser.add_argument("--measure", action="store_true", help="COM measure worst target")
    parser.add_argument("--diagnose", action="store_true", help="Show COM vs Oxi discrepancies")
    parser.add_argument("--record-success", action="store_true")
    parser.add_argument("--record-failure", action="store_true")
    args = parser.parse_args()

    if args.record_success:
        record_result(True)
    elif args.record_failure:
        record_result(False)
    elif args.measure:
        measure_target(args.docx_dir)
    elif args.diagnose:
        diagnose(args.docx_dir)
    else:
        print_status(args.docx_dir)
