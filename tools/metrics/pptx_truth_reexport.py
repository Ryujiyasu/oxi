# -*- coding: utf-8 -*-
"""Re-export every PowerPoint truth PDF, one PowerPoint PROCESS per deck.

Why a process per deck: an embedded family PowerPoint loads for one deck stays
mapped after `Close` (gen_pptx_embedonly: a later deck with no embedded Poppins
drew in Calibri), and `pptx_truth_audit.refresh` runs every deck through one
`PowerPoint.Application`, so each deck's truth depended on the decks exported
before it. Cloud fonts also arrive on demand, so the stored truths were aimed
at whatever the machine held on the day they were made (blind 26: stored truth
= cloud Montserrat Regular, fresh export = the deck's embedded Medium).

Sets (each keeps the previous PDF beside it as `<name>.pdf.<stamp>.bak`):
    dev     pptx_benchmark/dev/pptx/<stem>.pptx -> dev/pdf/<stem>.pdf
            (and clears dev/ppt_png/<stem>/ so scoring re-rasters it)
    blind   pptx_benchmark/pptx/<NN>__*.pptx    -> ssim_pptx/ppt_pdf/<NN>.pdf
    newdev  pptx_newdev/pptx/<stem>.pptx        -> pptx_newdev/pdf/<stem>.pdf

The CloudFonts inventory at export time is written next to the log so a later
truth can be compared against the font state that produced it.

    python tools/metrics/pptx_truth_reexport.py --sets dev,blind,newdev --stamp 20261118

NEVER run while a renderer gate is producing PNGs.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
BENCH = REPO / "pipeline_data" / "pptx_benchmark"
NEWDEV = REPO / "pipeline_data" / "pptx_newdev"
CLOUD = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "FontCache" / "4" / "CloudFonts"


def jobs(sets: list[str]) -> list[tuple[str, Path, Path, Path | None]]:
    out = []
    if "dev" in sets:
        for p in sorted((BENCH / "dev" / "pptx").glob("*.pptx")):
            out.append(("dev", p, BENCH / "dev" / "pdf" / f"{p.stem}.pdf", BENCH / "dev" / "ppt_png" / p.stem))
    if "blind" in sets:
        for p in sorted((BENCH / "pptx").glob("*.pptx")):
            pdf = BENCH / "ssim_pptx" / "ppt_pdf" / f"{p.stem[:2]}.pdf"
            if pdf.exists():
                out.append(("blind", p, pdf, None))
    if "newdev" in sets:
        for p in sorted((NEWDEV / "pptx").glob("*.pptx")):
            out.append(("newdev", p, NEWDEV / "pdf" / f"{p.stem}.pdf", None))
    return out


def inventory() -> dict[str, list[str]]:
    if not CLOUD.is_dir():
        return {}
    return {d.name: sorted(f.name for f in d.iterdir()) for d in sorted(CLOUD.iterdir()) if d.is_dir()}


def export_one(pptx: Path, pdf: Path) -> None:
    import pythoncom
    import win32com.client
    subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
    time.sleep(1.5)
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(pptx.resolve()), WithWindow=False)
        try:
            tmp = pdf.with_suffix(".pdf.tmp.pdf")
            pres.SaveAs(str(tmp), 32)  # 32 = ppSaveAsPDF
        finally:
            pres.Close()
    finally:
        app.Quit()
    tmp.replace(pdf)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--sets", default="dev,blind,newdev")
    ap.add_argument("--stamp", required=True)
    ap.add_argument("--only", default="", help="comma list of stem prefixes")
    a = ap.parse_args()
    sets = a.sets.split(",")
    only = [s for s in a.only.split(",") if s]
    log_dir = REPO / "pipeline_data" / "pptx_truth_reexport"
    log_dir.mkdir(parents=True, exist_ok=True)
    (log_dir / f"cloudfonts_{a.stamp}.json").write_text(json.dumps(inventory(), indent=1), encoding="utf-8")
    todo = [j for j in jobs(sets) if not only or any(j[1].stem.startswith(o) for o in only)]
    print(f"{len(todo)} decks", flush=True)
    failed = []
    for n, (kind, pptx, pdf, cache) in enumerate(todo, 1):
        bak = pdf.with_name(pdf.name + f".{a.stamp}.bak")
        if pdf.is_file() and not bak.exists():
            shutil.copy2(pdf, bak)
        t0 = time.time()
        try:
            export_one(pptx, pdf)
        except Exception as e:  # noqa: BLE001
            failed.append(pptx.name)
            print(f"[{n}/{len(todo)}] FAIL {kind} {pptx.stem[:50]}: {e}", flush=True)
            continue
        if cache is not None and cache.is_dir():
            shutil.rmtree(cache)
        print(f"[{n}/{len(todo)}] {kind} {pptx.stem[:50]} {time.time() - t0:.0f}s", flush=True)
    print("failed:", failed, flush=True)


if __name__ == "__main__":
    main()
