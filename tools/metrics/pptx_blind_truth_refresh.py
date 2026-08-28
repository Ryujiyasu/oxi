# -*- coding: utf-8 -*-
"""Re-export the blind corpus's truth PDFs with PowerPoint COM.

The stored PDFs were exported 2026-08-09. Office downloads cloud fonts ON DEMAND,
so a deck's truth encodes whichever faces the machine happened to hold that day --
and `pptx_cloudfont_staleness` finds 20 blind decks that name a face which only
arrived afterwards. Measuring against those is aiming at a past state of the
machine, not at PowerPoint.

    python tools/metrics/pptx_blind_truth_refresh.py --docs 18,28
    python tools/metrics/pptx_blind_truth_refresh.py --stale

The previous PDF is kept as `<doc>.pdf.<stamp>.bak`, so any measurement can be
re-run against the old truth to see what the refresh moved.

★NEVER run this while the renderer is producing PNGs: a live PowerPoint COM
instance during a render run has corrupted whole decks (`pptx_com_render_must_not_overlap`).
★One cold open per deck, one session (`pptx_truth_pdf_first_open_is_cold`) --
which is what the original 2026-08-09 export did, at ~30s per deck.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
PDF_DIR = ROOT / "ssim_pptx" / "ppt_pdf"
OXI_PNG = ROOT / "ssim_pptx" / "oxi_png"


def stale_docs() -> list[str]:
    out = subprocess.run(
        [sys.executable, str(REPO / "tools" / "metrics" / "pptx_cloudfont_staleness.py")],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    ).stdout
    tail = [ln for ln in out.splitlines() if "decks whose truth PDF predates" in ln]
    if not tail:
        sys.exit("staleness audit produced no verdict line")
    return tail[0].split(": ")[1].split(",")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docs", default="")
    ap.add_argument("--stale", action="store_true")
    ap.add_argument("--stamp", default="20260809")
    args = ap.parse_args()

    docs = ([d.strip() for d in args.docs.split(",") if d.strip()]
            or (stale_docs() if args.stale else []))
    if not docs:
        sys.exit("need --docs or --stale")

    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    by_idx = {f"{i['idx']:02d}": ROOT / "pptx" / i["local"] for i in manifest}
    targets = []
    for doc in docs:
        key = f"{int(doc):02d}"
        src = by_idx.get(key)
        if src is None or not src.exists():
            sys.exit(f"no deck for {doc}")
        targets.append((key, src))

    import win32com.client
    print(f"refreshing {len(targets)} truth PDFs: {', '.join(k for k, _ in targets)}")
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for key, src in targets:
            pdf = PDF_DIR / f"{key}.pdf"
            if pdf.is_file():
                shutil.copy2(pdf, pdf.with_suffix(f".pdf.{args.stamp}.bak"))
            pres = app.Presentations.Open(str(src.resolve()), WithWindow=False)
            try:
                pres.SaveAs(str(pdf), 32)  # ppSaveAsPDF
            finally:
                pres.Close()
            shutil.rmtree(OXI_PNG / key, ignore_errors=True)
            print(f"  {key} refreshed (oxi_png cleared)", flush=True)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
