"""Scan baseline docx files for docGrid + Normal style rPr sz.

Finds candidates for verifying the empty-para/grid line-height formula:
  gap = (ceil(N/grid)*grid + N)/2,  N = round(fs * 1.2)

Outputs CSV-ish table of (name, grid_type, line_pitch_tw, normal_sz_pt, N_pred, grid_pt).
"""
import re, zipfile, json
from pathlib import Path

DOCX_DIR = Path("tools/golden-test/documents/docx")
BASELINE = json.load(open("pipeline_data/ssim_baseline.json"))

baseline_names = set(BASELINE.keys())

RE_DOCGRID = re.compile(
    r'<w:docGrid\b[^/]*?(?:w:type="(?P<type>[^"]+)")?[^/]*?(?:w:linePitch="(?P<pitch>\d+)")?[^/]*?/>'
)
RE_NORMAL_SZ = re.compile(
    r'<w:style[^>]+w:default="1"[^>]*>.*?<w:sz w:val="(\d+)"',
    re.DOTALL
)
RE_DEFAULT_SZ = re.compile(
    r'<w:docDefaults>.*?<w:sz w:val="(\d+)"',
    re.DOTALL
)
# NB: some docs put sz in rPrDefault rather than Normal style
RE_NORMAL_STYLE = re.compile(
    r'<w:style[^>]+w:styleId="([^"]+)"[^>]+w:default="1"',
)

def extract_info(docx_path: Path):
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            doc = z.read("word/document.xml").decode("utf-8", errors="replace")
            try:
                styles = z.read("word/styles.xml").decode("utf-8", errors="replace")
            except KeyError:
                styles = ""
    except Exception as e:
        return None

    # docGrid (first occurrence; all sections should match for simple case)
    m = RE_DOCGRID.search(doc)
    grid_type = None; pitch_tw = None
    if m:
        # Re-search with more flexible attribute order
        grid_el = m.group(0)
        tm = re.search(r'w:type="([^"]+)"', grid_el)
        pm = re.search(r'w:linePitch="(\d+)"', grid_el)
        grid_type = tm.group(1) if tm else None
        pitch_tw = int(pm.group(1)) if pm else None

    # Normal style size (first style with default=1)
    sz_hp = None
    for m in re.finditer(r'<w:style[^>]+w:default="1"[^>]*>.*?</w:style>', styles, re.DOTALL):
        blk = m.group(0)
        szm = re.search(r'<w:sz w:val="(\d+)"', blk)
        if szm:
            sz_hp = int(szm.group(1))
            break
    if sz_hp is None:
        # fallback: docDefaults rPrDefault
        szm = re.search(r'<w:rPrDefault>.*?<w:sz w:val="(\d+)"', styles, re.DOTALL)
        if szm:
            sz_hp = int(szm.group(1))

    return dict(
        grid_type=grid_type,
        pitch_tw=pitch_tw,
        sz_hp=sz_hp,
    )


def predict_gap(fs_pt, grid_pt):
    import math
    N = math.floor(fs_pt * 1.2 + 0.5)  # round half-up
    eff = math.ceil(N / grid_pt) * grid_pt
    return (eff + N) / 2.0, N


def main():
    rows = []
    for docx in sorted(DOCX_DIR.glob("*.docx")):
        # strip .docx → must be in baseline
        stem = docx.stem
        if stem not in baseline_names:
            continue
        info = extract_info(docx)
        if not info: continue
        rows.append((stem, info))

    # filter: docGrid lines mode, have a Normal sz, grid != natural_line
    candidates = []
    for name, info in rows:
        if info["grid_type"] != "lines": continue
        if info["sz_hp"] is None: continue
        if info["pitch_tw"] is None: continue
        fs = info["sz_hp"] / 2.0
        grid = info["pitch_tw"] / 20.0
        pred_gap, N = predict_gap(fs, grid)
        # interesting cases: where naive grid-full-snap would differ from formula
        naive = grid
        diff = abs(pred_gap - naive)
        candidates.append((name, fs, grid, N, pred_gap, naive, diff))

    print(f"Scanned {len(rows)} docs, {len(candidates)} with docGrid lines + Normal sz")
    print(f"{'name':60s} {'fs':>5s} {'grid':>5s} {'N':>3s} {'gap':>6s} {'naive':>6s} {'Δ':>5s}")
    print("-" * 100)
    for r in sorted(candidates, key=lambda x: -x[6]):
        name, fs, grid, N, gap, naive, diff = r
        print(f"{name:60s} {fs:>5.1f} {grid:>5.1f} {N:>3d} {gap:>6.2f} {naive:>6.2f} {diff:>5.2f}")


if __name__ == "__main__":
    main()
