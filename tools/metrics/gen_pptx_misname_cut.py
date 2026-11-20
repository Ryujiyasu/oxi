# -*- coding: utf-8 -*-
"""Probe cut from a real deck (d51): which face draws "Josefin Sans SemiBold"?

Synthetic decks (gen_pptx_misname.py) never load their embedded parts, so each
arm here is d51 itself with one change. d51 files four typefaces; its
"Josefin Sans SemiBold" slots are byte copies of its "Josefin Sans" slots.

  A  unchanged                     (control: truth draws the cloud SemiBold)
  B  "Josefin Sans" entry dropped  (runs naming it -> "Noto Sans", so nothing
     asks the cloud for Josefin Sans): the SemiBold slots are no longer copies
  C  "Josefin Sans SemiBold" -> "Zzz Probe SemiBold" everywhere (no machine or
     cloud face of that name): does PowerPoint load the copy at all?

CloudFonts is listed before and after; a new file means PowerPoint downloaded
something and the arm's result (and later truths) must be read with that.

    python tools/metrics/gen_pptx_misname_cut.py
"""
from __future__ import annotations

import glob
import os
import re
import subprocess
import sys
import time
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[2]
SRC = Path(glob.glob(str(ROOT / "pipeline_data" / "pptx_benchmark" / "dev" / "pptx" / "d51*.pptx"))[0])
OUT = ROOT / "pipeline_data" / "pptx_probes" / "misname_cut"
CLOUD = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "FontCache" / "4" / "CloudFonts"


def cloud_list() -> set[str]:
    return {str(p.relative_to(CLOUD)) for p in CLOUD.rglob("*") if p.is_file()}


def rewrite(dst: Path, fn) -> None:
    zin = zipfile.ZipFile(SRC)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.endswith(".xml"):
            data = fn(item.filename, data.decode("utf8")).encode("utf8")
        zout.writestr(item, data)
    zout.close()


def arm_b(name: str, s: str) -> str:
    if name == "ppt/presentation.xml":
        s = re.sub(r'<p:embeddedFont><p:font typeface="Josefin Sans"/>.*?</p:embeddedFont>', "", s, flags=re.S)
    return s.replace('typeface="Josefin Sans"', 'typeface="Noto Sans"')


def arm_c(name: str, s: str) -> str:
    return s.replace('typeface="Josefin Sans SemiBold"', 'typeface="Zzz Probe SemiBold"')


def export(src: Path) -> Path:
    import win32com.client

    pdf = src.with_suffix(".pdf")
    subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
    time.sleep(3)
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        p = app.Presentations.Open(os.path.normpath(src), ReadOnly=True, WithWindow=False)
        p.SaveAs(os.path.normpath(pdf), 32)
        p.Close()
    finally:
        app.Quit()
    return pdf


def report(pdf: Path) -> None:
    import pymupdf

    d = pymupdf.open(pdf)
    for b in d[0].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                if s["text"].strip():
                    print(f"    {s['font']:28s} {s['size']:6.2f} {s['text'][:40]!r}")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    arms = {"A_unchanged": lambda n, s: s, "B_nocopy": arm_b, "C_zzz": arm_c}
    for name, fn in arms.items():
        before = cloud_list()
        path = OUT / f"d51_{name}.pptx"
        rewrite(path, fn)
        pdf = export(path)
        new = cloud_list() - before
        print(f"{name}  (CloudFonts new: {sorted(new) or 'none'})")
        report(pdf)


if __name__ == "__main__":
    main()
