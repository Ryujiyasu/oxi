"""Side-by-side Oxi-vs-Libra-vs-Word report across the 4 comparison axes:
  - Feature coverage (open rate from libra_open_report.json)
  - Pagination correctness (Phase 1 gate)
  - LLA (line layout agreement)
  - SSIM (pixel-level)

Pre-req: render_libra.py + compute_ssim_libra.py + measure_pagination_libra.py
         + pagination_diff_libra.py + compare_lla_libra_vs_oxi.py have all
         been run.

Reads:
  pipeline_data/libra_open_report.json
  pipeline_data/pagination_diff/_summary.json         (Oxi-side Phase 1 gate)
  pipeline_data/pagination_diff_libra/_summary.json   (Libra-side)
  pipeline_data/lla_compare/_summary.json
  pipeline_data/libra_vs_oxi_ssim.json

Writes:
  pipeline_data/oxi_vs_libra_report.md   (human-readable summary)
  pipeline_data/oxi_vs_libra_report.json (machine-readable details)

Usage:
    python tools/metrics/oxi_vs_libra_report.py
"""
from __future__ import annotations

import json
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PD = REPO_ROOT / "pipeline_data"

SOURCES = {
    "open": PD / "libra_open_report.json",
    "pagination_oxi": PD / "pagination_diff" / "_summary.json",
    "pagination_libra": PD / "pagination_diff_libra" / "_summary.json",
    "lla": PD / "lla_compare" / "_summary.json",
    "ssim": PD / "libra_vs_oxi_ssim.json",
}


def load(path: Path) -> dict | None:
    if not path.is_file():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def fmt_pct(x: float | None) -> str:
    return "n/a" if x is None else f"{x*100:.1f}%"


def fmt_delta_pct(x: float | None) -> str:
    if x is None:
        return "n/a"
    sign = "+" if x >= 0 else ""
    return f"{sign}{x*100:.1f}%"


def fmt_score(x: float | None) -> str:
    return "n/a" if x is None else f"{x:.4f}"


def fmt_delta_score(x: float | None) -> str:
    if x is None:
        return "n/a"
    sign = "+" if x >= 0 else ""
    return f"{sign}{x:.4f}"


def main():
    data = {k: load(p) for k, p in SOURCES.items()}
    missing = [k for k, v in data.items() if v is None]
    if missing:
        print(f"# WARN: missing sources: {missing}")

    rows = []

    # --- Feature coverage (open rate) ---
    open_data = data.get("open")
    if open_data:
        rows.append({
            "axis": "Feature coverage (open rate)",
            "oxi": "n/a (Oxi opens all baseline by definition; Libra coverage check)",
            "libra": f"{open_data['n_convert_ok']}/{open_data['n_total']} converted, "
                     f"{open_data['n_rasterize_ok']}/{open_data['n_total']} rasterized "
                     f"({open_data['convert_rate']*100:.1f}% convert, "
                     f"{open_data['rasterize_rate']*100:.1f}% rasterize)",
            "winner": "—",
        })

    # --- Pagination Phase 1 gate ---
    pag_libra = data.get("pagination_libra")
    if pag_libra:
        mean_l = pag_libra.get("mean_libra_score")
        mean_o = pag_libra.get("mean_oxi_score_in_join")
        delta = pag_libra.get("mean_delta_libra_minus_oxi")
        oxi_pass_in_join = pag_libra.get("n_oxi_pass_in_join")
        libra_pass = pag_libra["n_libra_pass"]
        n_total = pag_libra["n_total"]
        winner = "Oxi" if (delta is not None and delta < -0.01) else ("Libra" if (delta is not None and delta > 0.01) else "≈ tied")
        rows.append({
            "axis": "Pagination Phase 1 gate (mean score over joined docs)",
            "oxi": f"{oxi_pass_in_join}/{n_total} pass, mean {fmt_score(mean_o)}" if mean_o is not None else "n/a",
            "libra": f"{libra_pass}/{n_total} pass, mean {fmt_score(mean_l)}",
            "winner": f"{winner} ({fmt_delta_score(delta)})",
        })

    # --- LLA ---
    lla = data.get("lla")
    if lla:
        s = lla.get("summary", {})
        mean_o = s.get("mean_oxi_rate")
        mean_l = s.get("mean_libra_rate")
        delta = s.get("mean_delta")
        n_lib = s.get("n_libra_better", 0)
        n_oxi = s.get("n_oxi_better", 0)
        n_tied = s.get("n_tied", 0)
        winner = "Oxi" if (delta is not None and delta < -0.005) else ("Libra" if (delta is not None and delta > 0.005) else "≈ tied")
        rows.append({
            "axis": "LLA (LCS line agreement, mean across docs)",
            "oxi": f"mean {fmt_pct(mean_o)}",
            "libra": f"mean {fmt_pct(mean_l)}",
            "winner": f"{winner} ({fmt_delta_pct(delta)}; Oxi+:{n_oxi}, Libra+:{n_lib}, tied:{n_tied})",
        })

    # --- SSIM ---
    ssim_data = data.get("ssim")
    if ssim_data:
        s = ssim_data.get("summary", {})
        mean_o = s.get("mean_oxi_score")
        mean_l = s.get("mean_libra_score")
        delta = s.get("mean_delta_libra_minus_oxi")
        n_lib = s.get("n_libra_better", 0)
        n_oxi = s.get("n_oxi_better", 0)
        n_tied = s.get("n_within_001", 0)
        winner = "Oxi" if (delta is not None and delta < -0.005) else ("Libra" if (delta is not None and delta > 0.005) else "≈ tied")
        rows.append({
            "axis": "SSIM (Word PNG vs target PNG, mean across pages)",
            "oxi": f"mean {fmt_score(mean_o)} ({s.get('n_scored', 0)} pages)",
            "libra": f"mean {fmt_score(mean_l)}",
            "winner": f"{winner} ({fmt_delta_score(delta)}; Oxi+:{n_oxi}, Libra+:{n_lib}, tied:{n_tied})",
        })

    # --- Print table ---
    import sys
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print()
    print("=" * 100)
    print(" Oxi vs LibreOffice -- fidelity to Word reference render")
    print("=" * 100)
    for row in rows:
        print()
        print(f"  Axis:   {row['axis']}")
        print(f"  Oxi:    {row['oxi']}")
        print(f"  Libra:  {row['libra']}")
        print(f"  Winner: {row['winner']}")
    print()
    print("=" * 100)

    # --- Markdown report ---
    md_lines = [
        "# Oxi vs LibreOffice — Fidelity Comparison",
        "",
        "Reference: Microsoft Word (current rendered output is the ground truth).",
        "",
        "| Axis | Oxi | LibreOffice | Winner |",
        "|------|-----|-------------|--------|",
    ]
    for row in rows:
        md_lines.append(
            f"| {row['axis']} | {row['oxi']} | {row['libra']} | {row['winner']} |"
        )
    md_lines.append("")
    md_lines.append("## Detail sources")
    for k, p in SOURCES.items():
        md_lines.append(f"- `{k}`: `{p.relative_to(REPO_ROOT)}`" + ("" if p.is_file() else "  **(missing)**"))

    md_path = PD / "oxi_vs_libra_report.md"
    md_path.write_text("\n".join(md_lines), encoding="utf-8")
    print(f"# wrote {md_path}")

    json_path = PD / "oxi_vs_libra_report.json"
    json_path.write_text(json.dumps({"rows": rows, "sources_loaded": {k: data[k] is not None for k in SOURCES}}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"# wrote {json_path}")


if __name__ == "__main__":
    main()
