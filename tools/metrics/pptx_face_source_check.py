# -*- coding: utf-8 -*-
"""Does S-FACECOLLIDE retrodict which copy each truth PDF actually used?

For every embedded part the collision probe calls TAKEN *as of the truth PDF's
export date*, PowerPoint should have drawn the cache's copy; for every part it
calls free, the deck's own. A PDF subset keeps the source font's `name` table, so
name ID 5 says which -- provided the two disagree on version, which is the only
case this can judge.

    python tools/metrics/pptx_face_source_check.py --asof 2026-08-09

Reports agree / disagree / undecidable (the two copies share a version string).
"""
from __future__ import annotations

import argparse
import json
import os
import struct
import subprocess
import sys
import zipfile
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
CLOUD = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "FontCache" / "4" / "CloudFonts"
PROBE = REPO / "tools" / "metrics" / "pptx_part_name_collision.py"
SFNT = (bytes([0, 1, 0, 0]), b"OTTO", b"true")


def sfnt_names(blob: bytes) -> dict[int, str]:
    if blob[:4] not in SFNT:
        return {}
    tbl = None
    for i in range(struct.unpack_from(">H", blob, 4)[0]):
        rec = 12 + 16 * i
        if blob[rec:rec + 4] == b"name":
            off, ln = struct.unpack_from(">II", blob, rec + 8)
            tbl = blob[off:off + ln]
            break
    if not tbl:
        return {}
    n, so = struct.unpack_from(">HH", tbl, 2)
    out: dict[int, str] = {}
    for i in range(n):
        pid, _e, _l, nid, ln, off = struct.unpack_from(">HHHHHH", tbl, 6 + 12 * i)
        raw = tbl[so + off:so + off + ln]
        try:
            s = raw.decode("utf-16-be" if pid == 3 else "latin-1").strip("\x00").strip()
        except UnicodeDecodeError:
            continue
        if s:
            out.setdefault(nid, s)
    return out


def eot_version(blob: bytes) -> str:
    off = 16 + 10 + 2 + 4 + 2 + 2 + 16 + 8 + 4 + 16
    out = []
    for _ in range(3):
        off += 2
        n = struct.unpack_from("<H", blob, off)[0]
        off += 2
        out.append(blob[off:off + n].decode("utf-16-le", "replace"))
        off += n
    return out[2]


def cache_versions() -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    for path in CLOUD.rglob("*"):
        if path.suffix.lower() not in (".ttf", ".otf"):
            continue
        nm = sfnt_names(path.read_bytes())
        fam = nm.get(16) or nm.get(1)
        if fam and nm.get(5):
            out.setdefault(fam.lower(), set()).add(nm[5])
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--asof", default="2026-08-09")
    ap.add_argument("--docs", default="")
    args = ap.parse_args()

    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    docs = ([d.strip() for d in args.docs.split(",") if d.strip()]
            or [f"{i['idx']:02d}" for i in manifest
                if (ROOT / "ssim_pptx" / "ppt_pdf" / f"{i['idx']:02d}.pdf").exists()])
    verdicts = json.loads(subprocess.run(
        [sys.executable, str(PROBE), ",".join(docs), f"--asof={args.asof}", "--json"],
        capture_output=True, text=True, encoding="utf-8", check=True).stdout)

    cache = cache_versions()
    agree = disagree = undecidable = 0
    for doc in docs:
        src = next(ROOT / "pptx" / i["local"] for i in manifest if f"{i['idx']:02d}" == doc)
        pdf = pymupdf.open(ROOT / "ssim_pptx" / "ppt_pdf" / f"{doc}.pdf")
        seen, in_pdf = set(), set()
        for pno in range(len(pdf)):
            for xref, _e, _t, _b, _r, _en in pdf.get_page_fonts(pno):
                if xref in seen:
                    continue
                seen.add(xref)
                nm = sfnt_names(pdf.extract_font(xref)[3])
                if nm.get(5):
                    in_pdf.add(nm[5])
        pdf.close()
        with zipfile.ZipFile(src) as z:
            for part, verdict in verdicts.get(doc, {}).items():
                blob = z.read(f"ppt/fonts/{part}")
                own = eot_version(blob)
                fams = [f for f in cache if own not in cache[f]]
                # Only the family this part declares can answer for it.
                from_eot = eot_version  # noqa: F841  (kept for readability)
                theirs = set()
                for fam, vers in cache.items():
                    if fam in part.lower().replace("-", " ") or fam.replace(" ", "") in part.lower():
                        theirs |= vers
                if not theirs or theirs == {own}:
                    undecidable += 1
                    continue
                want_cache = verdict == "TAKEN"
                saw_cache = bool(theirs & in_pdf)
                saw_own = own in in_pdf
                if want_cache and saw_cache and not saw_own:
                    agree += 1
                elif not want_cache and saw_own:
                    agree += 1
                elif want_cache and saw_own and not saw_cache:
                    disagree += 1
                    print(f"  {doc} {part}: predicted CACHE, PDF has the part ({own})")
                elif not want_cache and saw_cache and not saw_own:
                    disagree += 1
                    print(f"  {doc} {part}: predicted PART, PDF has the cache")
                else:
                    undecidable += 1
    total = agree + disagree
    print(f"\nas of {args.asof}: {agree} agree, {disagree} disagree "
          f"({agree / total * 100:.1f}% of {total} decidable parts; "
          f"{undecidable} undecidable)")


if __name__ == "__main__":
    main()
