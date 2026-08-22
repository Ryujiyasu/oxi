# -*- coding: utf-8 -*-
"""Export the cloud-font probe -- and re-export a real deck -- with PowerPoint.

The probe answers "which of these families does PowerPoint resolve today".
`--deck d19` additionally re-exports a CORPUS deck to a scratch PDF, which
answers the separate question of whether that deck's stored truth PDF is still
reproducible: d19's truth is Calibri throughout although it asks for Nunito, and
if a fresh export now draws Nunito then the truth is a snapshot of a machine
state, not a property of the deck.

NEVER run this while the renderer is producing PNGs -- a live PowerPoint COM
instance during a render run has corrupted whole decks before.

Usage:
    python tools/metrics/export_pptx_cloudfont.py
    python tools/metrics/export_pptx_cloudfont.py --deck d19 --deck d06
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
PROBE = REPO / "pipeline_data" / "pptx_probes" / "cloudfont" / "probe_cloudfont.pptx"
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev" / "pptx"
RE_EXPORT = REPO / "pipeline_data" / "pptx_probes" / "cloudfont" / "reexport"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--deck", action="append", default=[],
                    help="corpus deck id (d06) to re-export to a scratch PDF")
    args = ap.parse_args()

    jobs: list[tuple[Path, Path]] = [(PROBE, PROBE.with_suffix(".pdf"))]
    if args.deck:
        RE_EXPORT.mkdir(parents=True, exist_ok=True)
        for want in args.deck:
            hits = sorted(DEV.glob(f"{want}__*.pptx"))
            if not hits:
                sys.exit(f"no deck matched {want}")
            jobs.append((hits[0], RE_EXPORT / f"{want}.pdf"))

    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for src, out in jobs:
            pres = app.Presentations.Open(str(src.resolve()), WithWindow=False)
            try:
                pres.SaveAs(str(out), 32)  # 32 = ppSaveAsPDF
            finally:
                pres.Close()
            print("wrote", out)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
