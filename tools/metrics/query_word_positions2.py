"""Fast Word COM line position query - MoveDown approach."""
import os, json, time
import win32com.client as win32

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(SCRIPT_DIR, "docx_tests_isolated")
FONT_METRICS = os.path.join(SCRIPT_DIR, "..", "..", "crates", "oxidocs-core", "src", "font", "data", "font_metrics.json")

FM_MAP = {
    "yumin": "Yu Mincho Regular", "yugothic": "Yu Gothic Regular",
    "century": "Century", "tnr": "Times New Roman",
    "calibri": "Calibri", "arial": "Arial",
    "msmincho": "MS Mincho", "msgothic": "MS Gothic",
}

TESTS = [
    ("single_nogrid", [
        ("arial", 10.5), ("arial", 11), ("arial", 12), ("arial", 14),
        ("calibri", 10.5), ("calibri", 11), ("calibri", 12), ("calibri", 14),
        ("century", 10.5), ("century", 11), ("century", 12), ("century", 14),
        ("tnr", 10.5), ("tnr", 11), ("tnr", 12), ("tnr", 14),
        ("msgothic", 10.5), ("msgothic", 11), ("msgothic", 12), ("msgothic", 14),
        ("msmincho", 10.5), ("msmincho", 11), ("msmincho", 12), ("msmincho", 14),
        ("yugothic", 10.5), ("yugothic", 11), ("yugothic", 12), ("yugothic", 14),
        ("yumin", 10.5), ("yumin", 11), ("yumin", 12), ("yumin", 14),
    ]),
]


def measure_file(word, filepath):
    doc = word.Documents.Open(filepath)
    sel = word.Selection
    sel.HomeKey(Unit=6)  # wdStory

    positions = []
    for _ in range(15):
        y = sel.Information(6)  # wdVerticalPositionRelativeToPage
        positions.append(y)
        moved = sel.MoveDown(Unit=5, Count=1)  # wdLine
        if moved == 0:
            break
        new_y = sel.Information(6)
        if abs(new_y - y) < 0.01:
            break

    doc.Close(0)

    gaps = []
    for i in range(1, len(positions)):
        g = positions[i] - positions[i-1]
        if 0 < g < 100:
            gaps.append(g)

    if not gaps:
        return None, positions
    gaps.sort()
    return gaps[len(gaps) // 2], positions


def main():
    with open(FONT_METRICS, encoding="utf-8") as f:
        font_metrics = {fm["family"]: fm for fm in json.load(f)}

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    results = []

    for mode, tests in TESTS:
        print(f"{'Font':<12} {'Size':>5} {'COM_gap':>8} {'win_h':>7} {'ratio':>7}")
        print("-" * 50)

        for fid, size in tests:
            fname = f"{mode}_{fid}_{size}pt.docx"
            fpath = os.path.abspath(os.path.join(INPUT_DIR, fname))

            gap, pos = measure_file(word, fpath)
            fm_key = FM_MAP.get(fid, "")
            fm = font_metrics.get(fm_key, {})
            upm = fm.get("units_per_em", 1)
            win_h = (fm.get("win_ascent", 0) + fm.get("win_descent", 0)) / upm * size

            if gap:
                print(f"  {fid:<12} {size:>5} {gap:>8.4f} {win_h:>7.3f} {gap/size:>7.4f}")
                results.append({"font_id": fid, "size": size, "gap": round(gap, 4),
                                "win_h": round(win_h, 4), "positions": [round(p, 4) for p in pos]})
            else:
                print(f"  {fid:<12} {size:>5}  NO GAPS  pos={pos[:4]}")

    word.Quit()

    out = os.path.join(SCRIPT_DIR, "output", "word_com_positions.json")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w") as f:
        json.dump(results, f, indent=2)
    print(f"\nSaved to {out}")


if __name__ == "__main__":
    main()
