"""Day 33 part 68 (R7.21 a47e6 campaign) — Measure nested-table atLeast
row precision across minimal repros.

For each variant (N01..N11):
1. Word COM: open and capture each nested-table row's start Y position.
2. Oxi: render with OXI_DUMP_TABLE=1 and parse the entry_cursor_y values.
3. Compare per-row delta (Oxi cy - Word y).
4. Report which variant shows the +0.5pt/row drift (matches a47e6 row 6).

Output: pipeline_data/nested_atleast_results.csv
        Per-row Word_y / Oxi_cy / delta for every variant.

Run: python tools/metrics/measure_nested_atleast_repro.py
"""

from __future__ import annotations
import os
import re
import sys
import csv
import time
import glob
import subprocess
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
REPRO_DIR = os.path.join(REPO, "tools", "golden-test", "repros", "nested_atleast")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")
OUT_CSV = os.path.join(REPO, "pipeline_data", "nested_atleast_results.csv")


def measure_word(path: str) -> list[dict]:
    """Open docx in Word, return list of {row_idx, page, y} for the nested
    table's rows (identified by their text content 'R{r}c0')."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    rows = []
    try:
        doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
        time.sleep(0.3)

        # Iterate paragraphs; nested cell paragraphs contain "R1c0", "R2c0", etc.
        # We capture the first occurrence of each "R{r}c0" as the row's start.
        seen = set()
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            text = p.Range.Text.replace("\r", "").replace("\x07", "").strip()
            m = re.fullmatch(r"R(\d+)c0", text)
            if m:
                r_idx = int(m.group(1))
                if r_idx in seen:
                    continue
                seen.add(r_idx)
                start_rng = doc.Range(p.Range.Start, p.Range.Start)
                rows.append({
                    "row_idx": r_idx,
                    "page": start_rng.Information(3),
                    "y": round(start_rng.Information(6), 3),
                })

        doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    return sorted(rows, key=lambda r: r["row_idx"])


def measure_oxi(path: str) -> list[dict]:
    """Render with OXI_DUMP_TABLE=1 and parse nested-table entry_cursor_y.

    The PARENT table is row=0 (1 row only) with cells. The NESTED table
    is rendered inside parent's cell, so its TBL_DUMP entries appear AFTER
    the parent's row=0 entry. We capture nested table's row=0..N entries.

    Strategy: find first 'row=0 entry_cursor_y' (parent), then subsequent
    'row=0..N entry_cursor_y' that are the nested table's rows."""
    env = os.environ.copy()
    env["OXI_DUMP_TABLE"] = "1"
    proc = subprocess.run(
        [GDI, os.path.abspath(path), os.path.join(REPO, "nul")],
        env=env,
        capture_output=True,
        text=False,
    )
    text = proc.stderr.decode("utf-8", errors="replace")

    # Clean up renderer's PNG artifacts
    for f in os.listdir(REPO):
        if f.startswith("nul_p") and f.endswith(".png"):
            try:
                os.remove(os.path.join(REPO, f))
            except OSError:
                pass

    # Parse all entry_cursor_y lines.
    entries = []
    for line in text.splitlines():
        m = re.search(
            r"row=(\d+) entry_cursor_y=([\d.]+) row_height_pre=([\d.]+).*?n_cells=(\d+)",
            line,
        )
        if m:
            entries.append({
                "row_idx_dump": int(m.group(1)),
                "cy": float(m.group(2)),
                "rh_pre": float(m.group(3)),
                "n_cells": int(m.group(4)),
            })

    # The first entry is the PARENT table row 0. Subsequent entries form
    # the nested table. Nested table rows are identified by n_cells=2
    # (the nested tables have 2 cells per row).
    if not entries:
        return []
    parent_y = entries[0]["cy"]

    # Find nested rows: those after parent with n_cells matching nested
    # structure (2 cells for col=2 nested, 1 cell for col=1 etc.).
    nested = entries[1:]
    return nested


def main() -> int:
    docxes = sorted(glob.glob(os.path.join(REPRO_DIR, "N*.docx")))
    if not docxes:
        print(f"No repros found in {REPRO_DIR}")
        return 1

    all_results = []
    for docx in docxes:
        name = os.path.splitext(os.path.basename(docx))[0]
        print(f"\n=== {name} ===")
        try:
            word_rows = measure_word(docx)
            oxi_rows = measure_oxi(docx)
        except Exception as e:
            print(f"  [NG] error: {e}")
            continue

        print(f"  Word: {len(word_rows)} rows, Oxi: {len(oxi_rows)} rows")

        # Pair by row index (1-based for Word, 0-based for Oxi)
        n = min(len(word_rows), len(oxi_rows))
        for i in range(n):
            w_r = word_rows[i]
            o_r = oxi_rows[i]
            delta = o_r["cy"] - w_r["y"]
            print(
                f"  row {i+1}: word_y={w_r['y']:.2f} (pg {w_r['page']}) "
                f"oxi_cy={o_r['cy']:.2f} rh_pre={o_r['rh_pre']:.2f} "
                f"Δ={delta:+.2f}"
            )
            all_results.append({
                "variant": name,
                "row": i + 1,
                "word_page": w_r["page"],
                "word_y": w_r["y"],
                "oxi_cy": o_r["cy"],
                "oxi_rh_pre": o_r["rh_pre"],
                "delta": round(delta, 3),
            })

    # Write CSV
    os.makedirs(os.path.dirname(OUT_CSV), exist_ok=True)
    if all_results:
        with open(OUT_CSV, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(all_results[0].keys()))
            w.writeheader()
            w.writerows(all_results)
        print(f"\nCSV: {OUT_CSV}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
