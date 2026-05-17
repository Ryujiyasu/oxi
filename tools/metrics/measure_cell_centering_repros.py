"""COM-measure Word's y reporting for cell_centering_repro/ minimal repros.

For each repro, measure each paragraph:
- y_pg = Information(6) wdVerticalPositionRelativeToPage at collapsed start
- font size, line spacing rule (int), spacing value
- in_table flag

Then run Oxi via oxi-gdi-renderer with --dump-layout and capture first
element's y per para_idx. Compute dy = oxi_y - word_y_pg.

For each repro print expected vs observed dy and identify the mismatch.

Run from repo root:
    python tools/metrics/measure_cell_centering_repros.py
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path

import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO_ROOT = Path(__file__).parent.parent.parent
REPRO_DIR = Path(__file__).parent / "cell_centering_repro"
OUT_JSON = Path(__file__).parent / "cell_centering_measurements.json"
OXI_GDI = REPO_ROOT / "tools" / "oxi-gdi-renderer" / "target" / "release" / "oxi-gdi-renderer.exe"

WD_VERT_PAGE = 6
WD_VERT_TEXT = 7
WD_HORIZ_PAGE = 5
WD_PAGE = 3


def measure_word(docx_path: Path) -> list[dict]:
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
    time.sleep(0.25)
    out = []
    try:
        n = doc.Paragraphs.Count
        for pi in range(1, n + 1):
            p = doc.Paragraphs(pi)
            rng = p.Range
            text = (rng.Text or "")[:30]
            start_rng = doc.Range(rng.Start, rng.Start)
            rec = {"i": pi, "text": text}
            try:
                rec["page"] = start_rng.Information(WD_PAGE)
            except Exception:
                rec["page"] = None
            try:
                rec["y_pg"] = float(start_rng.Information(WD_VERT_PAGE))
            except Exception:
                rec["y_pg"] = None
            try:
                rec["x_pg"] = float(start_rng.Information(WD_HORIZ_PAGE))
            except Exception:
                rec["x_pg"] = None
            try:
                rec["in_table"] = bool(rng.Information(12))
            except Exception:
                rec["in_table"] = None
            fmt = p.Format
            rec["line_spacing"] = float(fmt.LineSpacing)
            rec["line_spacing_rule"] = int(fmt.LineSpacingRule)
            try:
                rec["fs"] = float(p.Range.Characters(1).Font.Size)
                rec["font"] = (p.Range.Characters(1).Font.NameFarEast or p.Range.Characters(1).Font.Name)
            except Exception:
                rec["fs"] = None
                rec["font"] = None
            out.append(rec)
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()
    return out


def measure_oxi(docx_path: Path) -> list[dict]:
    """Run oxi-gdi-renderer --dump-layout and return first element per para_idx."""
    out_prefix = str(docx_path.parent / f"_oxi_{docx_path.stem}")
    layout_json = str(docx_path.parent / f"_oxi_{docx_path.stem}_layout.json")
    result = subprocess.run(
        [str(OXI_GDI), str(docx_path), out_prefix, f"--dump-layout={layout_json}"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(f"  [WARN] oxi-gdi-renderer failed for {docx_path.name}: {result.stderr[:200]}")
        return []
    # oxi-gdi-renderer writes to Temp path occasionally; check both
    candidates = [
        Path(layout_json),
        Path(os.path.expandvars(rf"%LOCALAPPDATA%\Temp\_oxi_{docx_path.stem}_layout.json")),
    ]
    layout_path = next((p for p in candidates if p.exists()), None)
    if not layout_path:
        print(f"  [WARN] layout dump not found for {docx_path.name}")
        return []
    layout = json.loads(layout_path.read_text(encoding="utf-8"))
    per_para = {}
    for page_idx, page in enumerate(layout.get("pages", [])):
        for el in page.get("elements", []):
            if el.get("type") != "text":
                continue
            pi = el.get("para_idx")
            if pi is None:
                continue
            # Keep first element per (page, para_idx) — first text fragment is glyph-top
            key = pi
            if key not in per_para:
                per_para[key] = {
                    "para_idx": pi,
                    "page": page_idx + 1,
                    "y": el["y"],
                    "h": el["h"],
                    "x": el["x"],
                    "font_size": el.get("font_size"),
                    "in_table": el.get("cell_row_idx") is not None,
                }
    return [per_para[k] for k in sorted(per_para.keys())]


def main():
    results = {}
    docx_files = sorted(REPRO_DIR.glob("*.docx"))
    for dx in docx_files:
        print(f"\n=== {dx.name} ===")
        w = measure_word(dx)
        o = measure_oxi(dx)
        # Pair: Word i=k corresponds to Oxi para_idx=k-1
        rows = []
        for wp in w:
            oxi_entry = next((e for e in o if e["para_idx"] == wp["i"] - 1), None)
            if oxi_entry:
                dy = oxi_entry["y"] - wp["y_pg"]
            else:
                dy = None
            rows.append({
                "i": wp["i"],
                "in_table": wp["in_table"],
                "rule": wp["line_spacing_rule"],
                "line_spacing": wp["line_spacing"],
                "fs": wp["fs"],
                "font": wp["font"],
                "word_y_pg": wp["y_pg"],
                "oxi_y": oxi_entry["y"] if oxi_entry else None,
                "oxi_h": oxi_entry["h"] if oxi_entry else None,
                "dy": dy,
                "text": wp["text"],
            })
        results[dx.stem] = rows
        # Print summary
        for r in rows[:6]:
            in_t = "T" if r["in_table"] else "B"
            oxi_y_str = "----" if r["oxi_y"] is None else f"{r['oxi_y']:>6.2f}"
            dy_str = " ----" if r["dy"] is None else f"{r['dy']:>+5.2f}"
            oxi_h_str = "----" if r["oxi_h"] is None else f"{r['oxi_h']:>5.2f}"
            print(
                f"  pi={r['i']:<2} {in_t} rule={r['rule']} fs={r['fs']:.1f} "
                f"word_y={r['word_y_pg']:>6.2f} "
                f"oxi_y={oxi_y_str} "
                f"dy={dy_str} "
                f"oxi_h={oxi_h_str} {r['text']}"
            )

    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nWrote {OUT_JSON}")


if __name__ == "__main__":
    main()
