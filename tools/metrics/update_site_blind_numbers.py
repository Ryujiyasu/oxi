# -*- coding: utf-8 -*-
"""Regenerate the blind-benchmark scatter plots in docs/index.html and
docs/ja/index.html from the current SSIM result files.

Only the data inside each scatter <svg> changes: the <circle> points and the
"Oxi ahead on N/50" caption. Axes, labels and layout are left untouched, so the
plots stay visually identical apart from the measured values.

  python update_site_blind_numbers.py           # rewrite both pages
  python update_site_blind_numbers.py --check   # report, write nothing
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]

RESULTS = {
    "en": REPO / "pipeline_data/en_benchmark/ssim_blindB50/_result.json",
    "ja": REPO / "pipeline_data/ja_benchmark/ssim_blind50/_result.json",
}
PAGES = [REPO / "docs/index.html", REPO / "docs/ja/index.html"]

# aria-label engine name -> result-file key
ENGINES = {
    "LibreOffice": "libre",
    "ONLYOFFICE": "oo",
    "SILURUS": "silurus",
    "BetterOffice": "betteroffice",
    "eigenpal": "eigenpal",
}

X0, X1, Y0, Y1 = 42.0, 328.0, 300.0, 26.0
LO, HI = 0.4, 1.0


def point(other: float, oxi: float) -> tuple[float, float]:
    cx = X0 + (min(max(other, LO), HI) - LO) / (HI - LO) * (X1 - X0)
    cy = Y0 - (min(max(oxi, LO), HI) - LO) / (HI - LO) * (Y0 - Y1)
    return round(cx, 1), round(cy, 1)


def circles(rows: list[dict], eng: str) -> tuple[str, int, int]:
    out, wins, n = [], 0, 0
    for r in rows:
        a = r.get("oxi", {}).get("common_mean")
        b = r.get(eng, {}).get("common_mean")
        if a is None or b is None:
            continue
        n += 1
        if a > b + 0.0005:
            wins += 1
        cx, cy = point(b, a)
        label = f'{r["type"]}: Oxi {a:.3f} / {{eng}} {b:.3f}'
        out.append(f'  <circle cx="{cx}" cy="{cy}" r="3.4" fill="#264653" '
                   f'fill-opacity="0.55"><title>{label}</title></circle>')
    return "\n".join(out), wins, n


def rewrite(html: str, rows_by_lang: dict[str, list[dict]]) -> tuple[str, list[str]]:
    notes: list[str] = []
    # each scatter svg starts at its aria-label and ends at </svg>
    # EN page: 'Scatter plot of 50 blind[ Japanese] documents: Oxi SSIM vs <Eng> SSIM'
    # JA page: '散布図: [初見の日本語 ]50 文書...の Oxi と <Eng> の SSIM 比較'
    pattern = re.compile(
        r'(<svg viewBox="0 0 340 340" role="img" aria-label="'
        r'(?:Scatter plot of 50 blind(?P<ja_en> Japanese)? documents: Oxi SSIM vs (?P<eng_en>\w+) SSIM'
        r'|散布図: (?P<ja_jp>初見の日本語 )?50 文書(?:における)?の? ?Oxi と (?P<eng_jp>\w+) の SSIM 比較)'
        r'[^"]*"[^>]*>)(.*?)(</svg>)',
        re.S)

    seen = {"count": 0}

    def repl(m: re.Match) -> str:
        head = m.group(1)
        body, tail = m.group(len(m.groups()) - 1), m.group(len(m.groups()))
        eng_name = m.group("eng_en") or m.group("eng_jp")
        ja_flag = m.group("ja_en") or m.group("ja_jp")
        # the JA page's last two labels omit the "Japanese" marker; the page
        # order is 5 English scatters then 5 Japanese ones, so fall back to
        # position when the label itself is ambiguous.
        seen["count"] += 1
        lang = "ja" if (ja_flag or seen["count"] > 5) else "en"
        eng = ENGINES.get(eng_name)
        if eng is None:
            notes.append(f"  ?? unknown engine in aria-label: {eng_name}")
            return m.group(0)
        rows = rows_by_lang[lang]
        pts, wins, n = circles(rows, eng)
        pts = pts.replace("{eng}", eng_name)
        # drop the old circles, keep everything else
        body_wo = re.sub(r'\n?  <circle [^\n]*\n?', "\n", body)
        body_wo = re.sub(r"\n{2,}", "\n", body_wo)
        # re-insert the points just before the caption text line
        cap = re.search(r'(  <text x="42" y="14"[^\n]*\n)', body_wo)
        if not cap:
            notes.append(f"  ?? no caption line for {lang}/{eng_name}")
            return m.group(0)
        newcap = re.sub(r'(?:ahead on|優位) \d+/\d+', (f'ahead on {wins}/{n}' if 'ahead on' in cap.group(1) else f'優位 {wins}/{n}'), cap.group(1))
        body_new = body_wo[:cap.start(1)] + pts + "\n" + newcap + body_wo[cap.end(1):]
        notes.append(f"  {lang} vs {eng_name:12s}: {n} pts, Oxi ahead {wins}")
        return head + body_new + tail

    return pattern.sub(repl, html), notes


def main() -> None:
    check = "--check" in sys.argv
    rows_by_lang = {k: json.loads(p.read_text(encoding="utf-8"))["docs"]
                    for k, p in RESULTS.items()}
    for lang, rows in rows_by_lang.items():
        o = [r["oxi"]["common_mean"] for r in rows if r["oxi"]["common_mean"] is not None]
        print(f"{lang}: {len(rows)} docs, oxi mean {sum(o)/len(o):.4f}")
    for page in PAGES:
        html = page.read_text(encoding="utf-8")
        new, notes = rewrite(html, rows_by_lang)
        print(f"--- {page.relative_to(REPO)}")
        for n in notes:
            print(n)
        if not check and new != html:
            page.write_text(new, encoding="utf-8")
            print("  written")
        elif new == html:
            print("  (unchanged)")


if __name__ == "__main__":
    main()
