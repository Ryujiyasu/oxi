"""Measure every paragraph's start Y position in Word for db9ca18.

Cross-compare with Oxi BR_DUMP cursor_y to identify the +5pt drift source
between pi=19 and pi=36 (where Oxi gains over Word, causing +2pt overflow
at pi=36).

Output CSV: word_i, word_page, word_y_collapsed, text_prefix.
Then matches with Oxi cy via subprocess BR_DUMP capture.

Run: python tools/metrics/measure_db9ca18_per_para_y.py
"""

import sys
import os
import re
import csv
import time
import subprocess
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DOCX_PATH = os.path.join(
    REPO_ROOT,
    "tools",
    "golden-test",
    "documents",
    "docx",
    "db9ca18368cd_20241122_resource_open_data_01.docx",
)
GDI_RENDERER = os.path.join(
    REPO_ROOT,
    "tools",
    "oxi-gdi-renderer",
    "target",
    "release",
    "oxi-gdi-renderer.exe",
)


def measure_word_per_para() -> list[dict]:
    """Open db9ca18 and capture every paragraph's collapsed start Y via COM."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    results = []
    try:
        doc = word.Documents.Open(DOCX_PATH, ReadOnly=True)
        time.sleep(0.3)

        n = doc.Paragraphs.Count
        print(f"Word total paragraphs: {n}")
        for i in range(1, n + 1):
            p = doc.Paragraphs(i)
            rng = p.Range
            # Collapsed start range (R30 fix: avoid wdActiveEndPage of multi-page para)
            start_rng = doc.Range(rng.Start, rng.Start)
            page = start_rng.Information(3)
            y = start_rng.Information(6)
            text = p.Range.Text.replace("\r", "").replace("\x07", "")
            results.append({
                "word_i": i,
                "word_page": page,
                "word_y": round(y, 3),
                "text": text[:50],
            })
        doc.Close(SaveChanges=False)
    finally:
        word.Quit()
    return results


def measure_oxi_per_para() -> list[dict]:
    """Run oxi-gdi-renderer with OXI_DUMP_BREAK=1, parse BR_DUMP lines."""
    env = os.environ.copy()
    env["OXI_DUMP_BREAK"] = "1"
    proc = subprocess.run(
        [GDI_RENDERER, DOCX_PATH, os.path.join(REPO_ROOT, "nul")],
        env=env,
        capture_output=True,
        text=False,
    )
    text = proc.stderr.decode("utf-8", errors="replace")
    results = []
    for line in text.splitlines():
        m = re.search(
            r"\[BR_DUMP\] pi=(\d+) line0 cursor_y=([\d.]+).*?brk=(\w+)\s+text=\"([^\"]*)\"",
            line,
        )
        if m:
            pi = int(m.group(1))
            cy = float(m.group(2))
            brk = m.group(3) == "true"
            ptext = m.group(4)[:50]
            results.append({"pi": pi, "cy": round(cy, 3), "brk": brk, "text": ptext})
    # Clean up the renderer's accidental nul_p*.png files
    for f in os.listdir(REPO_ROOT):
        if f.startswith("nul_p") and f.endswith(".png"):
            try:
                os.remove(os.path.join(REPO_ROOT, f))
            except OSError:
                pass
    return results


def main() -> int:
    print("Measuring Word per-paragraph Y...")
    word_data = measure_word_per_para()
    print(f"Got {len(word_data)} Word paragraphs\n")

    print("Measuring Oxi per-paragraph cursor_y...")
    oxi_data = measure_oxi_per_para()
    print(f"Got {len(oxi_data)} Oxi BR_DUMP entries\n")

    # Build text-to-data lookup for cross-matching.
    # Match Word i to Oxi pi by text prefix (first 15 non-whitespace chars).
    def normalize(s: str) -> str:
        s = s.replace("　", " ")
        s = re.sub(r"\s+", " ", s).strip()
        return s

    word_by_text: dict[str, dict] = {}
    for r in word_data:
        t = normalize(r["text"])
        if t and t not in word_by_text:
            word_by_text[t] = r

    matches: list[dict] = []
    for o in oxi_data:
        ot = normalize(o["text"])
        if not ot:
            continue
        # Search word entries by prefix match (longest)
        best = None
        for wt, w in word_by_text.items():
            n = min(len(wt), len(ot))
            if n < 5:
                continue
            if wt[:n] == ot[:n]:
                if best is None or len(wt) > len(best[0]):
                    best = (wt, w)
        if best is None:
            continue
        wt, wrec = best
        delta = o["cy"] - wrec["word_y"]
        matches.append({
            "word_i": wrec["word_i"],
            "word_page": wrec["word_page"],
            "word_y": wrec["word_y"],
            "oxi_pi": o["pi"],
            "oxi_cy": o["cy"],
            "delta_y": round(delta, 3),
            "text": wrec["text"][:40],
        })

    print(f"Matched {len(matches)} paragraphs\n")
    print(f"{'wi':>3} {'wpg':>3} {'word_y':>8} {'pi':>3} {'oxi_cy':>8} {'Δ':>8} text")
    print("-" * 80)
    prev_delta = None
    cumul_change = 0.0
    for m in matches:
        flag = ""
        if prev_delta is not None:
            d = m["delta_y"] - prev_delta
            cumul_change += d
            if abs(d) > 0.5:
                flag = f" Δchg={d:+.2f}"
                if abs(d) > 5:
                    flag += " BIG"
        print(
            f"{m['word_i']:>3} {m['word_page']:>3} {m['word_y']:>8.2f} "
            f"{m['oxi_pi']:>3} {m['oxi_cy']:>8.2f} {m['delta_y']:>+8.2f} "
            f"{m['text']}{flag}"
        )
        prev_delta = m["delta_y"]

    # Write CSV for further analysis
    csv_path = os.path.join(REPO_ROOT, "pipeline_data", "db9ca18_per_para_y.csv")
    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(matches[0].keys()))
        w.writeheader()
        w.writerows(matches)
    print(f"\nCSV: {csv_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
