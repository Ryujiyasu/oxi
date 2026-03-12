"""
Query Word COM for actual line positions and heights.
Uses Range.Information to get precise positions.
"""
import os
import json
import win32com.client as win32

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(SCRIPT_DIR, "docx_tests_isolated")
MANIFEST = os.path.join(INPUT_DIR, "manifest.json")
FONT_METRICS = os.path.join(SCRIPT_DIR, "..", "..", "crates", "oxidocs-core", "src", "font", "data", "font_metrics.json")

FM_MAP = {
    "yumin": "Yu Mincho Regular",
    "yugothic": "Yu Gothic Regular",
    "century": "Century",
    "tnr": "Times New Roman",
    "calibri": "Calibri",
    "arial": "Arial",
    "msmincho": "MS Mincho",
    "msgothic": "MS Gothic",
}

# wdInformation constants
wdHorizontalPositionRelativeToPage = 5
wdVerticalPositionRelativeToPage = 6
wdFirstCharacterLineNumber = 10
wdLineWidth = 11


def main():
    with open(MANIFEST, encoding="utf-8") as f:
        manifest = json.load(f)

    with open(FONT_METRICS, encoding="utf-8") as f:
        font_metrics = {fm["family"]: fm for fm in json.load(f)}

    print("Starting Word...")
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    results = []
    # Test all 4 modes
    modes_to_check = ["single_nogrid", "115_nogrid", "single_grid", "default"]

    for mode in modes_to_check:
        entries = [e for e in manifest if e["mode"] == mode]
        print(f"\n=== Mode: {mode} ===")

        for entry in entries:
            path = os.path.join(INPUT_DIR, entry["filename"])
            fid = entry["font_id"]
            size = entry["size_pt"]

            try:
                doc = word.Documents.Open(os.path.abspath(path))

                # Get all lines' vertical positions by scanning character by character
                content = doc.Content
                total_chars = content.End
                lines = []
                prev_line_num = -1

                for char_idx in range(0, min(total_chars, 500)):
                    rng = doc.Range(char_idx, char_idx + 1)
                    text = rng.Text

                    if text in ['\r', '\n', '\x07']:
                        continue

                    try:
                        y_pos = rng.Information(wdVerticalPositionRelativeToPage)
                        line_num = rng.Information(wdFirstCharacterLineNumber)
                    except:
                        continue

                    if line_num != prev_line_num and line_num > 0:
                        lines.append({
                            "line": line_num,
                            "y": y_pos,
                            "char": text,
                        })
                        prev_line_num = line_num

                doc.Close(0)

                if len(lines) >= 2:
                    gaps = []
                    for i in range(1, len(lines)):
                        gap = lines[i]["y"] - lines[i-1]["y"]
                        if 0 < gap < 100:
                            gaps.append(gap)

                    if gaps:
                        gaps.sort()
                        median = gaps[len(gaps) // 2]
                        first_y = lines[0]["y"]

                        fm_key = FM_MAP.get(fid)
                        win_h = 0
                        if fm_key and fm_key in font_metrics:
                            fm = font_metrics[fm_key]
                            upm = fm["units_per_em"]
                            win_h = (fm["win_ascent"] + fm["win_descent"]) / upm * size

                        line_info = [f'{l["line"]}:{l["y"]:.2f}' for l in lines[:6]]
                        print(f"  {fid:<12} {size:>5}pt: first_y={first_y:.3f} gap={median:.3f}pt "
                              f"win_h={win_h:.3f} ratio={median/size:.4f} "
                              f"lines={line_info}")

                        results.append({
                            "mode": mode,
                            "font_id": fid,
                            "size_pt": size,
                            "first_y": round(first_y, 3),
                            "baseline_gap": round(median, 3),
                            "line_count": len(lines),
                            "win_height": round(win_h, 3),
                            "gap_ratio": round(median / size, 4),
                        })

            except Exception as e:
                print(f"  {fid} {size}pt: FAILED: {e}")

    word.Quit()

    out = os.path.join(SCRIPT_DIR, "output", "word_positions.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved to {out}")


if __name__ == "__main__":
    main()
