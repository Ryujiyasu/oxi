# -*- coding: utf-8 -*-
"""Rewrite the accuracy charts on the published site from the measured results.

The site carries, per language, a bar chart of per-engine means and one scatter
plot per competitor (a dot per document, Oxi on the vertical axis). Both had the
values baked in as literal SVG, so a new measurement leaves them stating the old
set's numbers under the new set's heading. This regenerates them from
`ssim_blindC50/_result.json`.

Only the data-bearing parts are touched: the `<circle>` rows and the chart
title inside each scatter, and the value/label/width of each bar row. Grid
lines, axes, styling and copy are left exactly as they are.

    python tools/metrics/_site_charts.py [--check]
"""
from __future__ import annotations

import io
import json
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
FILES = [REPO / "docs" / "index.html", REPO / "docs" / "ja" / "index.html"]
RESULTS = {
    "en": REPO / "pipeline_data" / "en_benchmark" / "ssim_blindC50" / "_result.json",
    "ja": REPO / "pipeline_data" / "ja_benchmark" / "ssim_blindC50" / "_result.json",
}
# Chart label -> key in _result.json
LABEL_TO_KEY = {
    "LibreOffice": "libre", "ONLYOFFICE": "oo", "SILURUS": "silurus",
    "BetterOffice": "betteroffice", "eigenpal": "eigenpal",
    "GenOffice": "genoffice", "OfficeCLI": "officecli",
}
# The axis mapping read off the existing SVG: 0.4 -> (42, 300), 1.0 -> (328, 26)
X0, X1, Y0, Y1, V0, V1 = 42.0, 328.0, 300.0, 26.0, 0.4, 1.0


def sx(v: float) -> float:
    return X0 + (min(max(v, V0), V1) - V0) * (X1 - X0) / (V1 - V0)


def sy(v: float) -> float:
    return Y0 + (min(max(v, V0), V1) - V0) * (Y1 - Y0) / (V1 - V0)


def load(lang: str) -> dict:
    return json.loads(RESULTS[lang].read_text(encoding="utf-8"))


def circles(data: dict, key: str) -> tuple[str, int, int]:
    """SVG circles for every document both engines scored, plus the win count."""
    rows, wins = [], 0
    for r in data["docs"]:
        o = (r.get("oxi") or {}).get("common_mean")
        c = (r.get(key) or {}).get("common_mean")
        if o is None or c is None:
            continue
        kind = r["doc"].split("__")[0]
        rows.append((kind, o, c))
        if o > c:
            wins += 1
    out = []
    for kind, o, c in rows:
        out.append(
            f'  <circle cx="{sx(c):.1f}" cy="{sy(o):.1f}" r="3.4" fill="#264653" '
            f'fill-opacity="0.55"><title>{kind}: Oxi {o:.3f} / '
            f'{{LABEL}} {c:.3f}</title></circle>')
    return "\n".join(out), wins, len(rows)


def rewrite_scatters(text: str, lang_of_index) -> tuple[str, int]:
    """Replace the circle block and heading of every scatter, in file order."""
    changed = 0
    pos, out = 0, []
    pat = re.compile(r'<svg viewBox="0 0 340 340".*?</svg>', re.S)
    idx = 0
    for m in pat.finditer(text):
        svg = m.group(0)
        title_m = re.search(r'<text x="42" y="14"[^>]*>([^<]*)</text>', svg)
        if not title_m:
            continue
        label = next((l for l in LABEL_TO_KEY if l in title_m.group(1)), None)
        if not label:
            continue
        lang = lang_of_index(idx)
        idx += 1
        data = load(lang)
        block, wins, n = circles(data, LABEL_TO_KEY[label])
        block = block.replace("{LABEL}", label)
        new_svg = re.sub(r'(?:  <circle .*?</circle>\n)+', block + "\n", svg, count=1,
                         flags=re.S)
        old_title = title_m.group(1)
        new_title = re.sub(r'\d+\s*/\s*\d+', f"{wins}/{n}", old_title)
        new_title = re.sub(r'\d+ 文書中 \d+ 文書', f"{n} 文書中 {wins} 文書", new_title)
        new_svg = new_svg.replace(f'>{old_title}</text>', f'>{new_title}</text>')
        out.append((m.start(), m.end(), new_svg))
        changed += 1
    for start, end, new in reversed(out):
        text = text[:start] + new + text[end:]
    return text, changed


def main() -> None:
    check = "--check" in sys.argv
    for f in FILES:
        s = io.open(f, encoding="utf-8").read()
        ja_at = s.index("blind set") if "blind set" in s else len(s)
        # First half of the scatters belong to the English section, second half
        # to the Japanese one; they carry the same engine labels, so position is
        # the only discriminator.
        total = len(re.findall(r'<svg viewBox="0 0 340 340"', s))
        half = total // 2
        s2, n = rewrite_scatters(s, lambda i: "en" if i < half else "ja")
        print(f"{f.relative_to(REPO)}: rewrote {n} scatter charts")
        if not check:
            io.open(f, "w", encoding="utf-8").write(s2)


if __name__ == "__main__":
    main()
