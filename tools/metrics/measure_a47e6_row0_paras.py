"""Day 33 part 68 R7.22 session 5 — Per-paragraph Y comparison within
a47e6 table 1 row 0 (the title/header row).

Row 0 over-pumps by +25pt (Word 262.5 vs Oxi 287.8). Walks each paragraph
inside row 0 cell, captures Word's actual rendered Y and Oxi's layout
JSON Y, computes per-paragraph delta.

Run: python tools/metrics/measure_a47e6_row0_paras.py
Output: pipeline_data/a47e6_row0_paras.csv
"""

from __future__ import annotations
import sys
import os
import json
import re
import csv
import time
import subprocess
import tempfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX = os.path.join(REPO, "tools", "golden-test", "documents", "docx",
                    "a47e6c6b2ca1_order_08.docx")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release", "oxi-gdi-renderer.exe")


def measure_word_row0_paras() -> list[dict]:
    """Word COM: get every paragraph in table 1 row 0, with its Y."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        doc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
        time.sleep(0.3)

        tbl = doc.Tables(1)
        tbl_start = tbl.Range.Start
        tbl_end = tbl.Range.End

        n_paras = doc.Paragraphs.Count
        para_seq = 0
        for i in range(1, n_paras + 1):
            p = doc.Paragraphs(i)
            p_start = p.Range.Start
            if p_start < tbl_start or p_start >= tbl_end:
                continue
            # Outer table 1 only
            try:
                nesting = p.Range.Tables.Count
            except Exception:
                nesting = 0
            if nesting != 1:
                continue
            # Only row 0 (RowIndex == 1 in COM 1-based)
            try:
                cell = p.Range.Cells(1)
                if cell.RowIndex != 1:
                    continue
            except Exception:
                continue

            start_rng = doc.Range(p_start, p_start)
            page = start_rng.Information(3)
            y = start_rng.Information(6)
            text = p.Range.Text.replace("\r", "").replace("\x07", "")
            para_seq += 1
            results.append({
                "para_seq": para_seq,
                "word_y": round(y, 3),
                "word_page": page,
                "text_len": len(text),
                "text_preview": repr(text[:30]),
            })
        doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    return results


def measure_oxi_row0_paras() -> list[dict]:
    """Render with --dump-layout and extract per-paragraph Y within row 0
    of table 1."""
    with tempfile.TemporaryDirectory(prefix="oxi_a47e6_") as tmp:
        out_prefix = os.path.join(tmp, "page_")
        dump_path = os.path.join(tmp, "layout.json")
        proc = subprocess.run(
            [GDI, os.path.abspath(DOCX), out_prefix, "150",
             f"--dump-layout={dump_path}"],
            capture_output=True,
            text=False,
        )
        if not os.path.exists(dump_path):
            print(f"[NG] dump not generated. stderr: {proc.stderr.decode('utf-8', errors='replace')[:500]}")
            return []
        with open(dump_path, "r", encoding="utf-8") as f:
            dump = json.load(f)

    # Find all text elements on page 1 with y within row 0's bounds
    # Row 0 in Oxi: entry_cursor_y=68.55, row_h=287.8 → y range [68.55, 356.35]
    # Group by paragraph (assume same y row = same line)
    elements = []
    for page in dump.get("pages", []):
        if page["page"] != 1:
            continue
        for el in page.get("elements", []):
            if el.get("type") != "text":
                continue
            y = el.get("y", 0)
            if 68.0 <= y <= 360.0:
                elements.append(el)

    # Cluster by y (0.5pt tolerance) to get unique line Y positions
    # then group by para_idx
    by_pi: dict = {}
    for el in elements:
        pi = el.get("para_idx")
        y = round(el["y"] * 2) / 2
        if pi is None:
            continue
        if pi not in by_pi:
            by_pi[pi] = {"first_y": y, "min_y": y, "max_y": y, "text_parts": []}
        slot = by_pi[pi]
        slot["min_y"] = min(slot["min_y"], y)
        slot["max_y"] = max(slot["max_y"], y)
        slot["text_parts"].append((el["y"], el["x"], el.get("text", "")))

    results = []
    seq = 0
    for pi in sorted(by_pi.keys()):
        slot = by_pi[pi]
        slot["text_parts"].sort(key=lambda yxt: (yxt[0], yxt[1]))
        text = "".join(t for _, _, t in slot["text_parts"])[:30]
        seq += 1
        results.append({
            "oxi_pi": pi,
            "oxi_seq": seq,
            "oxi_y": slot["min_y"],
            "oxi_text_preview": repr(text),
        })
    return results


def main() -> int:
    print("Measuring Word per-paragraph Y in row 0...")
    word_paras = measure_word_row0_paras()
    print(f"  {len(word_paras)} paragraphs found in row 0")

    print("\nMeasuring Oxi per-paragraph Y in row 0...")
    oxi_paras = measure_oxi_row0_paras()
    print(f"  {len(oxi_paras)} paragraphs found in Oxi page 1 row 0 area")

    # Pair by sequence
    print(f"\n{'seq':>3} {'word_y':>8} {'oxi_y':>8} {'delta':>8} {'len':>3} text")
    print("-" * 90)
    n = min(len(word_paras), len(oxi_paras))
    results = []
    prev_word_y = None
    prev_oxi_y = None
    for i in range(max(len(word_paras), len(oxi_paras))):
        w = word_paras[i] if i < len(word_paras) else {}
        o = oxi_paras[i] if i < len(oxi_paras) else {}
        word_y = w.get("word_y")
        oxi_y = o.get("oxi_y")
        if word_y is not None and oxi_y is not None:
            delta = oxi_y - word_y
            print(f"{i+1:>3} {word_y:>8.2f} {oxi_y:>8.2f} {delta:>+8.2f} "
                  f"{w.get('text_len',0):>3} W:{w.get('text_preview','')} O:{o.get('oxi_text_preview','')}")
            results.append({
                "seq": i + 1,
                "word_y": word_y,
                "oxi_y": oxi_y,
                "delta": round(delta, 3),
                "word_text": w.get("text_preview", ""),
                "oxi_text": o.get("oxi_text_preview", ""),
            })
        elif word_y is not None:
            print(f"{i+1:>3} {word_y:>8.2f} {'(none)':>8} {'-':>8} {w.get('text_len',0):>3} W:{w.get('text_preview','')}")
        else:
            print(f"{i+1:>3} {'(none)':>8} {oxi_y:>8.2f} {'-':>8} {'-':>3} O:{o.get('oxi_text_preview','')}")

    # Per-paragraph increments (= advance per paragraph)
    print(f"\nPer-paragraph INCREMENT (Word vs Oxi):")
    print(f"{'seq':>3} {'word_inc':>10} {'oxi_inc':>10} {'delta_inc':>10}")
    for i in range(1, len(results)):
        w_inc = results[i]["word_y"] - results[i-1]["word_y"]
        o_inc = results[i]["oxi_y"] - results[i-1]["oxi_y"]
        d_inc = o_inc - w_inc
        marker = " ★" if abs(d_inc) > 1 else ""
        print(f"{i+1:>3} {w_inc:>10.2f} {o_inc:>10.2f} {d_inc:>+10.2f}{marker}")

    out_csv = os.path.join(REPO, "pipeline_data", "a47e6_row0_paras.csv")
    os.makedirs(os.path.dirname(out_csv), exist_ok=True)
    if results:
        with open(out_csv, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(results[0].keys()))
            w.writeheader()
            w.writerows(results)
        print(f"\nCSV: {out_csv}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
