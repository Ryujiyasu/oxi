# -*- coding: utf-8 -*-
"""Which blind decks can S-CLOUDADV reach at all?

The cloud-advance path only substitutes numbers for a family the Office cloud
cache actually holds (`cloud_face_advances` returns None otherwise), and it is
consulted only from the break test. So a deck that names no cached family cannot
move, and does not need to be in the A/B -- which matters because a two-arm run
over 48 decks is a four-hour render.

A family counts as reachable when the cache holds a file whose `name` table
declares it (ID 16, else ID 1) -- the directory name is NOT the family
(`pptx_local_copy_beats_embedded`: the files are numeric and the directories are
close but not authoritative).

    python tools/metrics/pptx_cloudadv_census.py
"""
from __future__ import annotations

import json
import os
import re
import struct
import sys
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
CLOUD = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "FontCache" / "4" / "CloudFonts"
SFNT = (bytes([0, 1, 0, 0]), b"OTTO", b"true")


def sfnt_family(blob: bytes) -> str | None:
    if blob[:4] not in SFNT:
        return None
    n = struct.unpack(">H", blob[4:6])[0]
    name_off = name_len = 0
    for i in range(n):
        rec = 12 + 16 * i
        if blob[rec:rec + 4] == b"name":
            name_off, name_len = struct.unpack(">II", blob[rec + 8:rec + 16])
            break
    if not name_len:
        return None
    tbl = blob[name_off:name_off + name_len]
    count, str_off = struct.unpack(">HH", tbl[2:6])
    best = {}
    for i in range(count):
        rec = 6 + 12 * i
        pid, eid, lid, nid, ln, off = struct.unpack(">HHHHHH", tbl[rec:rec + 12])
        if nid not in (1, 16):
            continue
        raw = tbl[str_off + off:str_off + off + ln]
        try:
            txt = raw.decode("utf-16-be" if pid == 3 else "latin-1").strip("\x00").strip()
        except UnicodeDecodeError:
            continue
        if txt:
            best.setdefault(nid, txt)
    return best.get(16) or best.get(1)


def cloud_families() -> set[str]:
    out = set()
    for path in CLOUD.rglob("*"):
        if path.suffix.lower() not in (".ttf", ".otf"):
            continue
        try:
            fam = sfnt_family(path.read_bytes())
        except OSError:
            continue
        if fam:
            out.add(fam.lower())
    return out


def deck_typefaces(src: Path) -> set[str]:
    faces = set()
    with zipfile.ZipFile(src) as z:
        for info in z.infolist():
            if not info.filename.endswith(".xml"):
                continue
            if not info.filename.startswith(("ppt/slides/", "ppt/slideLayouts/",
                                             "ppt/slideMasters/", "ppt/theme/",
                                             "ppt/notesSlides/")):
                continue
            body = z.read(info).decode("utf-8", "replace")
            # Only <a:latin>. A theme's <a:cs>/<a:sym> name Mangal and Shruti on
            # every Office deck ever made, and those faces are reached only by a
            # complex-script run -- counting them makes the census say "all 48".
            faces.update(m.lower() for m in re.findall(r'<a:latin typeface="([^"]*)"', body) if m)
    return {f for f in faces if not f.startswith("+")}


def main() -> None:
    cloud = cloud_families()
    print(f"cache holds {len(cloud)} families\n")
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    hit = []
    for item in manifest:
        doc = f"{item['idx']:02d}"
        src = ROOT / "pptx" / item["local"]
        if not src.exists() or not (ROOT / "ssim_pptx" / "ppt_pdf" / f"{doc}.pdf").exists():
            continue
        faces = deck_typefaces(src)
        shared = sorted(faces & cloud)
        if shared:
            hit.append(doc)
            print(f"{doc}: {', '.join(shared)}")
    print(f"\n{len(hit)} decks reachable: {','.join(hit)}")


if __name__ == "__main__":
    main()
