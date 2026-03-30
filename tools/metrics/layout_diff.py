"""Layout DIFF: compare Word COM positions vs Oxi layout positions.
Identifies systematic font width/height errors and outputs corrections.

Usage: python layout_diff.py <input.docx>
"""
import win32com.client
import subprocess
import json
import sys
import os
from collections import defaultdict


def get_word_positions(docx_path):
    """Extract character positions from Word COM."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)

    ps = doc.PageSetup
    info = {
        "page_w": ps.PageWidth,
        "page_h": ps.PageHeight,
        "margin_l": ps.LeftMargin,
        "margin_r": ps.RightMargin,
        "margin_t": ps.TopMargin,
        "margin_b": ps.BottomMargin,
        "pages": doc.ComputeStatistics(2),
    }

    paragraphs = []
    for pi in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(pi)
        rng = p.Range
        align = p.Alignment  # 0=left,1=center,2=right,3=justify
        para_y = rng.Information(6)

        chars = rng.Characters
        n = chars.Count
        char_data = []
        prev_y = None
        line_idx = 0

        for ci in range(1, n + 1):
            try:
                c = chars(ci)
                ch = c.Text
                if ch in ('\r', '\x07'):
                    continue

                x = c.Information(5)
                y = c.Information(6)

                if prev_y is not None and abs(y - prev_y) > 1:
                    line_idx += 1
                prev_y = y

                char_data.append({
                    "ch": ch,
                    "x": round(x, 2),
                    "y": round(y, 2),
                    "line": line_idx,
                    "font": c.Font.Name,
                    "size": c.Font.Size,
                    "bold": bool(c.Font.Bold),
                })
            except:
                continue

        paragraphs.append({
            "index": pi,
            "align": align,
            "y": round(para_y, 2),
            "chars": char_data,
            "lines": line_idx + 1,
        })

    doc.Close(False)
    word.Quit()
    return {"info": info, "paragraphs": paragraphs}


def get_oxi_positions(docx_path):
    """Extract character positions from Oxi layout engine."""
    result = subprocess.run(
        ["cargo", "run", "--release", "--example", "layout_json", "--", docx_path],
        capture_output=True, text=True, errors="replace", cwd=os.path.dirname(os.path.abspath(__file__)) + "/../..",
        timeout=120,
    )

    paragraphs = []
    current_para = None
    prev_y = None
    line_idx = 0

    for line in result.stdout.strip().split("\n"):
        if line.startswith("PAGE\t"):
            continue
        if line.startswith("TEXT\t"):
            parts = line.split("\t")
            x, y, w, h_el = float(parts[1]), float(parts[2]), float(parts[3]), float(parts[4])
            fs = float(parts[5])
            font = parts[6]
            bold = parts[7] == "1"
        elif line.startswith("T\t"):
            text = line[2:]
            for ch in text:
                if current_para is None or (prev_y is not None and abs(y - prev_y) > h_el * 0.5):
                    if prev_y is not None and current_para and abs(y - prev_y) > h_el * 1.5:
                        paragraphs.append(current_para)
                        current_para = {"chars": [], "y": round(y, 2)}
                        line_idx = 0
                    elif current_para is None:
                        current_para = {"chars": [], "y": round(y, 2)}

                if prev_y is not None and abs(y - prev_y) > 1:
                    line_idx += 1
                prev_y = y

                current_para["chars"].append({
                    "ch": ch,
                    "x": round(x, 3),
                    "y": round(y, 3),
                    "line": line_idx,
                    "font": font,
                    "size": fs,
                    "bold": bold,
                    "width": round(w, 3),
                })
                x += w  # advance for next char in same text run
        else:
            continue

    if current_para and current_para["chars"]:
        paragraphs.append(current_para)

    return {"paragraphs": paragraphs}


def compare_and_analyze(word_data, oxi_data):
    """Compare positions and identify systematic errors."""
    # Flatten both to char lists
    word_chars = []
    for p in word_data["paragraphs"]:
        if p["align"] == 1:  # skip centered (COM positions unreliable)
            continue
        for c in p["chars"]:
            if c["ch"] not in ('\t', '\n', '\x0C', '\x0B', ' '):
                word_chars.append(c)

    oxi_chars = []
    for p in oxi_data["paragraphs"]:
        for c in p["chars"]:
            if c["ch"] not in ('\t', '\n', '\x0C', '\x0B', ' '):
                oxi_chars.append(c)

    # Match by character content (sequential matching)
    wi, oi = 0, 0
    matches = []
    while wi < len(word_chars) and oi < len(oxi_chars):
        wc = word_chars[wi]
        oc = oxi_chars[oi]
        if wc["ch"] == oc["ch"]:
            matches.append((wc, oc))
            wi += 1
            oi += 1
        elif wi + 1 < len(word_chars) and word_chars[wi + 1]["ch"] == oc["ch"]:
            wi += 1  # skip word char
        elif oi + 1 < len(oxi_chars) and oxi_chars[oi + 1]["ch"] == wc["ch"]:
            oi += 1  # skip oxi char
        else:
            wi += 1
            oi += 1

    # Analyze per-font/size statistics
    font_stats = defaultdict(lambda: {"dx": [], "dy": [], "count": 0})
    y_diffs = []

    for wc, oc in matches:
        dx = oc["x"] - wc["x"]
        dy = oc["y"] - wc["y"]
        key = f"{oc['font']}@{oc['size']}pt"
        font_stats[key]["dx"].append(dx)
        font_stats[key]["dy"].append(dy)
        font_stats[key]["count"] += 1
        y_diffs.append(dy)

    # Report
    print(f"\n{'='*60}")
    print(f"LAYOUT DIFF REPORT")
    print(f"{'='*60}")
    print(f"Matched: {len(matches)} chars")
    print(f"Word total: {len(word_chars)}, Oxi total: {len(oxi_chars)}")

    if y_diffs:
        import statistics
        print(f"\nY position (vertical):")
        print(f"  Mean dy: {statistics.mean(y_diffs):+.2f}pt")
        print(f"  Median dy: {statistics.median(y_diffs):+.2f}pt")
        print(f"  Stdev: {statistics.stdev(y_diffs):.2f}pt")

    print(f"\nPer-font width error (dx):")
    print(f"  {'Font':40s} {'N':>5} {'Mean dx':>9} {'Stdev':>7} {'Cumul/line':>11}")
    print(f"  {'-'*75}")

    corrections = {}
    for key in sorted(font_stats.keys()):
        s = font_stats[key]
        if s["count"] < 3:
            continue
        import statistics
        mean_dx = statistics.mean(s["dx"])
        stdev_dx = statistics.stdev(s["dx"]) if len(s["dx"]) > 1 else 0
        # Estimate cumulative error per 40-char line
        cumul = mean_dx  # dx is cumulative from line start
        print(f"  {key:40s} {s['count']:5d} {mean_dx:+9.3f} {stdev_dx:7.3f} {cumul:+11.3f}")

        # Extract font name and size for correction
        parts = key.rsplit("@", 1)
        font_name = parts[0]
        font_size = float(parts[1].replace("pt", ""))
        corrections[key] = {
            "font": font_name,
            "size": font_size,
            "mean_dx": round(mean_dx, 4),
            "stdev_dx": round(stdev_dx, 4),
            "samples": s["count"],
        }

    # Per-line Y comparison
    print(f"\nPer-line Y positions (first 20):")
    # Group by lines
    word_lines = defaultdict(list)
    oxi_lines = defaultdict(list)
    for wc, oc in matches:
        word_lines[wc["y"]].append(wc)
        oxi_lines[oc["y"]].append(oc)

    word_ys = sorted(word_lines.keys())
    oxi_ys = sorted(oxi_lines.keys())

    for i in range(min(20, len(word_ys), len(oxi_ys))):
        wy = word_ys[i]
        oy = oxi_ys[i]
        dy = oy - wy
        nw = len(word_lines[wy])
        no = len(oxi_lines[oy])
        marker = " ***" if abs(nw - no) > 2 else ""
        print(f"  Word y={wy:7.1f} ({nw:2d}ch)  Oxi y={oy:7.1f} ({no:2d}ch)  dy={dy:+6.2f}{marker}")

    return corrections


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python layout_diff.py <input.docx>")
        sys.exit(1)

    docx_path = sys.argv[1]
    print(f"Extracting Word positions...")
    word_data = get_word_positions(docx_path)
    print(f"  {sum(len(p['chars']) for p in word_data['paragraphs'])} chars from {len(word_data['paragraphs'])} paragraphs")

    print(f"Extracting Oxi positions...")
    oxi_data = get_oxi_positions(docx_path)
    print(f"  {sum(len(p['chars']) for p in oxi_data['paragraphs'])} chars from {len(oxi_data['paragraphs'])} paragraphs")

    corrections = compare_and_analyze(word_data, oxi_data)

    # Save corrections
    out_path = os.path.splitext(docx_path)[0] + "_layout_diff.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(corrections, f, indent=2, ensure_ascii=False)
    print(f"\nCorrections saved to: {out_path}")
