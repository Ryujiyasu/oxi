# -*- coding: utf-8 -*-
"""Which blind truth PDFs predate the cloud font PowerPoint would use today?

`pptx_truth_audit` asks the same question by NAME -- is the family among the
fonts the stored PDF names -- and that test is blind twice over. Decks exported
from Google Slides carry NUMERIC family names in their embedded parts (blind 18's
Montserrat subsets are both called "129"), so the name never matches; and the
Office cloud cache fills in ONE STYLE AT A TIME, so a family can be half present.

This asks it by CLOCK instead. A cache file that is NEWER than a deck's truth PDF
could not have been used when that PDF was exported, so any deck naming its family
is measured against a PowerPoint that had a different font in hand.

The split it was written to explain (2026-08-28), both truth PDFs exported
2026-08-09:

    28  Open Sans   all four styles cached 07-19   -> PowerPoint drew the CACHE
    18  Montserrat  only Regular cached 07-19;     -> PowerPoint drew the EMBEDDED
                    Bold arrived 08-18, and the        part, because the style it
                    deck needs Bold                    needed was not there yet

So the rule is per STYLE, not per family: PowerPoint takes the local copy when the
cache holds the face it needs, and the embedded part otherwise.

    python tools/metrics/pptx_cloudfont_staleness.py
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import struct
import sys
import zipfile
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
CLOUD = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "FontCache" / "4" / "CloudFonts"
SFNT = (bytes([0, 1, 0, 0]), b"OTTO", b"true")


def name_table(blob: bytes) -> dict[int, str]:
    if blob[:4] not in SFNT:
        return {}
    count = struct.unpack_from(">H", blob, 4)[0]
    tbl = None
    for i in range(count):
        rec = 12 + 16 * i
        if blob[rec:rec + 4] == b"name":
            off, ln = struct.unpack_from(">II", blob, rec + 8)
            tbl = blob[off:off + ln]
            break
    if not tbl:
        return {}
    n, str_off = struct.unpack_from(">HH", tbl, 2)
    out: dict[int, str] = {}
    for i in range(n):
        pid, _eid, _lid, nid, ln, off = struct.unpack_from(">HHHHHH", tbl, 6 + 12 * i)
        raw = tbl[str_off + off:str_off + off + ln]
        try:
            txt = raw.decode("utf-16-be" if pid == 3 else "latin-1").strip("\x00").strip()
        except UnicodeDecodeError:
            continue
        if txt:
            out.setdefault(nid, txt)
    return out


def cache_faces() -> dict[str, list[tuple[str, dt.datetime]]]:
    """family (lowercased) -> [(style, when it landed on this machine)]."""
    out: dict[str, list[tuple[str, dt.datetime]]] = {}
    for path in CLOUD.rglob("*"):
        if path.suffix.lower() not in (".ttf", ".otf"):
            continue
        try:
            nm = name_table(path.read_bytes())
        except OSError:
            continue
        fam = nm.get(16) or nm.get(1)
        if not fam:
            continue
        when = dt.datetime.fromtimestamp(path.stat().st_mtime)
        out.setdefault(fam.lower(), []).append((nm.get(2, "?"), when))
    return out


def deck_latin(src: Path) -> set[str]:
    faces: set[str] = set()
    with zipfile.ZipFile(src) as z:
        for info in z.infolist():
            if not info.filename.startswith(
                ("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/", "ppt/theme/")
            ) or not info.filename.endswith(".xml"):
                continue
            body = z.read(info).decode("utf-8", "replace")
            faces.update(m.lower() for m in re.findall(r'<a:latin typeface="([^"]*)"', body))
    return {f for f in faces if f and not f.startswith("+")}


def exported(pdf_path: Path) -> dt.datetime:
    with pymupdf.open(pdf_path) as d:
        raw = (d.metadata or {}).get("creationDate", "")
    m = re.match(r"D:(\d{14})", raw or "")
    if m:
        return dt.datetime.strptime(m.group(1), "%Y%m%d%H%M%S")
    return dt.datetime.fromtimestamp(pdf_path.stat().st_mtime)


def main() -> None:
    faces = cache_faces()
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    suspect = []
    for item in manifest:
        doc = f"{item['idx']:02d}"
        src = ROOT / "pptx" / item["local"]
        pdf_path = ROOT / "ssim_pptx" / "ppt_pdf" / f"{doc}.pdf"
        if not src.exists() or not pdf_path.exists():
            continue
        when = exported(pdf_path)
        late = [
            (fam, style, landed)
            for fam in sorted(deck_latin(src) & faces.keys())
            for style, landed in sorted(faces[fam])
            if landed > when
        ]
        if not late:
            continue
        suspect.append(doc)
        print(f"{doc}  truth exported {when:%Y-%m-%d %H:%M}")
        for fam, style, landed in late:
            print(f"      {fam} {style}  cached {landed:%Y-%m-%d %H:%M}")
    print(f"\n{len(suspect)} decks whose truth PDF predates a face they name: "
          f"{','.join(suspect)}")


if __name__ == "__main__":
    main()
