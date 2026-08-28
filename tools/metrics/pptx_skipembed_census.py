# -*- coding: utf-8 -*-
"""Which decks would `skipembed` touch, and through which font source?

`skipembed_on` drops an embedded part when the machine already resolves the
TYPEFACE the part is filed under -- and "the machine" is BOTH the installed fonts
and the Office cloud cache. The cloud half is the interesting one; the system half
is the dangerous one, because a deck that embeds a family this machine happens to
have installed (blind 47 embeds Caladea, which is installed here) loses its own
copy for the system's, and those are not always the same font.

This answers it without rendering: read `ppt/presentation.xml`'s
`p:embeddedFontLst`, and check each typeface against the system families and
against the cache's `name` tables.

    python tools/metrics/pptx_skipembed_census.py <system-font-list.txt>
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
    tbl = None
    for i in range(struct.unpack_from(">H", blob, 4)[0]):
        rec = 12 + 16 * i
        if blob[rec:rec + 4] == b"name":
            off, ln = struct.unpack_from(">II", blob, rec + 8)
            tbl = blob[off:off + ln]
            break
    if not tbl:
        return None
    n, so = struct.unpack_from(">HH", tbl, 2)
    best = {}
    for i in range(n):
        pid, _e, _l, nid, ln, off = struct.unpack_from(">HHHHHH", tbl, 6 + 12 * i)
        if nid not in (1, 16):
            continue
        raw = tbl[so + off:so + off + ln]
        try:
            s = raw.decode("utf-16-be" if pid == 3 else "latin-1").strip("\x00").strip()
        except UnicodeDecodeError:
            continue
        if s:
            best.setdefault(nid, s)
    return best.get(16) or best.get(1)


def main() -> None:
    system = {ln.strip().lower() for ln in Path(sys.argv[1]).read_text(
        encoding="utf-8", errors="replace").splitlines() if ln.strip()}
    cloud = set()
    for path in CLOUD.rglob("*"):
        if path.suffix.lower() in (".ttf", ".otf"):
            fam = sfnt_family(path.read_bytes())
            if fam:
                cloud.add(fam.lower())
    print(f"{len(system)} system families, {len(cloud)} cached families\n")

    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    by_sys, by_cloud = [], []
    for item in manifest:
        doc = f"{item['idx']:02d}"
        src = ROOT / "pptx" / item["local"]
        if not src.exists() or not (ROOT / "ssim_pptx" / "ppt_pdf" / f"{doc}.pdf").exists():
            continue
        with zipfile.ZipFile(src) as z:
            try:
                pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
            except KeyError:
                continue
        faces = {m.lower() for m in re.findall(r'<p:font typeface="([^"]+)"', pres)}
        s_hit = sorted(faces & system)
        c_hit = sorted((faces & cloud) - system)
        if not s_hit and not c_hit:
            continue
        if s_hit:
            by_sys.append(doc)
        if c_hit:
            by_cloud.append(doc)
        note = []
        if s_hit:
            note.append("SYSTEM: " + ", ".join(s_hit))
        if c_hit:
            note.append("cache: " + ", ".join(c_hit))
        print(f"{doc}  {'  |  '.join(note)}")
    print(f"\n{len(by_sys)} decks reached through an INSTALLED family: {','.join(by_sys)}")
    print(f"{len(by_cloud)} decks reached through the cache only: {','.join(by_cloud)}")


if __name__ == "__main__":
    main()
