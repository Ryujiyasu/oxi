# -*- coding: utf-8 -*-
"""Ask whether a stored reference PDF still reflects what PowerPoint does today.

A reference PDF is not permanent. Office downloads cloud fonts over time, and a
deck whose family was NOT in the cache on the day it was exported fell back to
the part the deck embeds; the same deck exported after the download uses the
cloud font instead, with different advances. The dev corpus already shows this
split cleanly (2026-08-26):

    deck  exported  family               PowerPoint used
    d10   08-23     Montserrat-Regular   local  (a=0.590, embedded is 0.607)
    d15   08-23     Barlow-Light         local  (a=0.506, embedded is 0.511)
    d35   08-10     OpenSans-Regular     EMBEDDED (a=0.55615, local is 0.5649)
    d08   08-10     Merriweather-Bold    EMBEDDED (a=0.566,  local is 0.561)

Four clean cases, and which one PowerPoint picked is decided entirely by WHICH
EXPORT SESSION the reference came from -- not by any property of the deck. Open
Sans and Merriweather are both in the cloud cache now, so if the split really is
staleness, re-exporting d35 and d08 today must flip them to the local font.
That is the test this script runs.

It re-exports to a SCRATCH path and never touches the stored reference: the
decision to refresh the corpus is the user's, because it moves every SSIM
number in the benchmark.

★Never run this while the renderer is producing PNGs
(`pptx-com-render-must-not-overlap`) -- PowerPoint and the renderer fight over
embedded-font resolution and the deck comes out wrong.

★One deck per COM session. `pptx-truth-pdf-first-open-is-cold`: a deck opened a
second time in one session resolves fonts differently, so each arm gets its own
PowerPoint.

Usage:
    python tools/metrics/pptx_reference_staleness.py            # d35, d08
    python tools/metrics/pptx_reference_staleness.py d35 d02
"""
from __future__ import annotations

import glob
import io
import subprocess
import sys
import tempfile
from pathlib import Path

import pymupdf
from fontTools.ttLib import TTFont

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

EXPORT_ONE = """
import sys
import win32com.client

src, out = sys.argv[1], sys.argv[2]
app = win32com.client.Dispatch("PowerPoint.Application")
try:
    pres = app.Presentations.Open(src, WithWindow=False)
    try:
        pres.SaveAs(out, 32)  # ppSaveAsPDF
    finally:
        pres.Close()
finally:
    app.Quit()
"""


def subsets(pdf: Path, pages: int = 30) -> dict[str, tuple[int | None, float | None]]:
    """PostScript name -> (usWeightClass, advance of 'a' in EM) for each subset."""
    doc = pymupdf.open(pdf)
    out: dict[str, tuple[int | None, float | None]] = {}
    for pno in range(min(len(doc), pages)):
        for xref, _, _, base, _, _, _ in doc[pno].get_fonts(full=True):
            try:
                _, _, _, buf = doc.extract_font(xref)
            except Exception:
                continue
            if not buf:
                continue
            try:
                t = TTFont(io.BytesIO(buf))
                ps = next((r.toUnicode() for r in t["name"].names if r.nameID == 6), base)
                cmap = t.getBestCmap()
                a = None
                if ord("a") in cmap:
                    a = round(t["hmtx"][cmap[ord("a")]][0] / t["head"].unitsPerEm, 5)
                out.setdefault(ps, (t["OS/2"].usWeightClass, a))
            except Exception:
                continue
    return out


def export(src: Path, out: Path) -> bool:
    """Export `src` in a PowerPoint of its own."""
    with tempfile.NamedTemporaryFile("w", suffix=".py", delete=False, encoding="utf-8") as fh:
        fh.write(EXPORT_ONE)
        driver = fh.name
    try:
        rc = subprocess.run(
            [sys.executable, driver, str(src.resolve()), str(out.resolve())],
            capture_output=True,
            text=True,
        )
        if rc.returncode != 0:
            print(f"  export failed: {rc.stderr.strip()[:200]}")
            return False
        return out.exists()
    finally:
        Path(driver).unlink(missing_ok=True)


def main() -> None:
    decks = sys.argv[1:] or ["d35", "d08"]
    scratch = DEV / "_staleness"
    scratch.mkdir(parents=True, exist_ok=True)
    for deck in decks:
        src = sorted(DEV.glob(f"pptx/{deck}__*.pptx"))
        ref = sorted(DEV.glob(f"pdf/{deck}__*.pdf"))
        if not src or not ref:
            print(f"{deck}: missing pptx or reference pdf")
            continue
        fresh = scratch / f"{deck}_today.pdf"
        print(f"\n=== {deck}  {src[0].name}")
        if not export(src[0], fresh):
            continue
        was, now = subsets(ref[0]), subsets(fresh)
        names = sorted(set(was) | set(now))
        print(f"{'PostScript name':<30}{'stored':>18}{'today':>18}   change")
        for n in names:
            w0, a0 = was.get(n, (None, None))
            w1, a1 = now.get(n, (None, None))
            if (w0, a0) == (w1, a1):
                continue
            s0 = "-" if a0 is None and w0 is None else f"w{w0} a={a0}"
            s1 = "-" if a1 is None and w1 is None else f"w{w1} a={a1}"
            note = "GONE" if n not in now else ("NEW" if n not in was else "CHANGED")
            print(f"{n:<30}{s0:>18}{s1:>18}   {note}")
        if was == now:
            print("  reference still current: identical font set")


if __name__ == "__main__":
    main()
