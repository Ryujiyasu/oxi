"""Measure empty-paragraph line height gap for the full G/H matrix.

For each .docx in empty_para_grid_repro/, open in Word via COM and measure
Y-position of para 1 and para 2 using Information(6) wdVerticalPositionRelativeToPage.
Gap = y2 - y1 = the empty paragraph's line height.

Outputs JSON: {name: {fs_pt, grid_pt, gap_pt, y1, y2}}.
"""
import os, json, time, re
from pathlib import Path
import win32com.client

REPRO_DIR = Path(__file__).parent / "empty_para_grid_repro"
OUT_JSON = Path(__file__).parent / "empty_para_grid_matrix.json"

NAME_RE = re.compile(r"^([GH]\d+)_sz(\d+)_(grid\d+|no_grid)")
# H40_norm21_pprrpr24, H41_norm21_pprrpr24_gothic, H42_norm21_pprrpr24_text,
# H43_norm21_pprrpr28, H44_norm21_pprrpr32, H50_norm21_text_runsz24, H51_norm21_text_runsz28
OVERRIDE_RE = re.compile(r"^H\d+_norm(\d+)_(?:text_runsz|pprrpr)(\d+)")
SPECIAL = {
    # Effective (fs, grid) for non-standard-named cases
    "H60_d77a_p3_exact": (12.0, 18.0),
    "H61_fe_layout_gothic": (12.0, 18.0),
    "H62_kern_gothic": (12.0, 18.0),
    "H63_szcs24_gothic": (12.0, 18.0),
    "H64_fe_mincho": (12.0, 18.0),
    "H65_fe_sz24": (12.0, 18.0),
    "H66_fe_sz21": (10.5, 18.0),
    "H67_d77a_p3_strip": (12.0, 18.0),
    "H68_no_szcs": (12.0, 18.0),
    "H69_no_bookmark": (12.0, 18.0),
}

def parse_name(stem: str):
    m = NAME_RE.match(stem)
    if m:
        sz_hp = int(m.group(2))
        grid_tag = m.group(3)
        fs_pt = sz_hp / 2.0
        if grid_tag == "no_grid":
            grid_pt = None
        else:
            tw = int(grid_tag[4:])
            grid_pt = tw / 20.0
        return fs_pt, grid_pt
    m = OVERRIDE_RE.match(stem)
    if m:
        # override case: effective para-mark fs comes from pPr.rPr or run rPr
        override_hp = int(m.group(2))
        return override_hp / 2.0, 18.0  # all H40-H51 use grid=360tw=18pt
    if stem in SPECIAL:
        return SPECIAL[stem]
    return None, None


def measure_one(word, path: Path):
    doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
    time.sleep(0.25)
    try:
        p1 = doc.Paragraphs(1).Range
        p2 = doc.Paragraphs(2).Range
        # Information(6) = wdVerticalPositionRelativeToPage
        y1 = doc.Range(p1.Start, p1.Start + 1).Information(6)
        y2 = doc.Range(p2.Start, p2.Start + 1).Information(6)
        gap = y2 - y1
        return y1, y2, gap
    finally:
        doc.Close(False)


def main():
    results = {}
    docs = sorted(REPRO_DIR.glob("*.docx"))
    print(f"Found {len(docs)} repro docs")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        for p in docs:
            stem = p.stem
            fs_pt, grid_pt = parse_name(stem)
            if fs_pt is None:
                print(f"skip (cant parse): {stem}")
                continue
            try:
                y1, y2, gap = measure_one(word, p)
                results[stem] = dict(fs_pt=fs_pt, grid_pt=grid_pt,
                                     y1=round(y1, 3), y2=round(y2, 3),
                                     gap_pt=round(gap, 3))
                grid_s = f"{grid_pt:>5.1f}" if grid_pt else "  --  "
                print(f"  {stem:30s} fs={fs_pt:>4.1f} grid={grid_s}  y1={y1:>6.2f}  y2={y2:>6.2f}  gap={gap:>6.3f}")
            except Exception as e:
                print(f"  {stem}: FAIL {e}")
                results[stem] = dict(fs_pt=fs_pt, grid_pt=grid_pt, error=str(e))
    finally:
        word.Quit()

    OUT_JSON.write_text(json.dumps(results, indent=2, ensure_ascii=False))
    print(f"\nwrote {OUT_JSON}")


if __name__ == "__main__":
    main()
