# -*- coding: utf-8 -*-
"""Is each deck's stored truth PDF still what PowerPoint would export today?

Office downloads cloud fonts ON DEMAND, so the set of families PowerPoint can
resolve GROWS over time. A truth PDF exported before a family arrived shows
PowerPoint's fallback (Calibri) where a fresh export shows the real face -- and
every SSIM number measured against it is then aimed at a state of the machine
rather than at PowerPoint.

`audit` finds those decks without touching PowerPoint: a family is suspect when
the deck asks for it, the family is CLOUD-ONLY (present in
`%LOCALAPPDATA%\\Microsoft\\FontCache\\4\\CloudFonts`, absent from the system, so
it can only have arrived by download) and it does NOT appear among the fonts the
stored PDF names.

`refresh` re-exports those decks with PowerPoint COM, keeping the previous PDF
next to it as `<deck>.pdf.<stamp>.bak`, and deletes the deck's cached reference
rasters so the next scoring run rebuilds them.

    python tools/metrics/pptx_truth_audit.py audit
    python tools/metrics/pptx_truth_audit.py refresh --decks d19,d24
    python tools/metrics/pptx_truth_audit.py refresh --all-suspect --stamp 20260823

NEVER run `refresh` while the renderer is producing PNGs -- a live PowerPoint COM
instance during a render run has corrupted whole decks before.
"""
from __future__ import annotations

import argparse
import re
import shutil
import struct
import sys
import zipfile
from collections import Counter
from pathlib import Path

import pymupdf

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
DEV = REPO / "pipeline_data" / "pptx_benchmark" / "dev"
PPTX_DIR = DEV / "pptx"
PDF_DIR = DEV / "pdf"
REF_ROOT = DEV / "ppt_png"
CLOUD_ROOT = Path(
    __import__("os").environ["LOCALAPPDATA"]
) / "Microsoft" / "FontCache" / "4" / "CloudFonts"


SFNT_TAGS = (bytes([0, 1, 0, 0]), b"OTTO", b"true")  # sfnt magic numbers


def norm(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", name.lower())


def font_family(path: Path) -> str | None:
    """The typographic family (name ID 16, else 1) out of a font's name table."""
    blob = path.read_bytes()
    if blob[:4] not in (b"\x00\x01\x00\x00", b"OTTO", b"true"):
        return None
    count = struct.unpack(">H", blob[4:6])[0]
    offset = None
    for index in range(count):
        record = 12 + 16 * index
        if blob[record:record + 4] == b"name":
            offset = struct.unpack(">I", blob[record + 8:record + 12])[0]
            break
    if offset is None:
        return None
    _fmt, entries, strings = struct.unpack(">HHH", blob[offset:offset + 6])
    best: dict[int, str] = {}
    for index in range(entries):
        record = offset + 6 + 12 * index
        pid, _eid, _lid, nid, length, off = struct.unpack(">HHHHHH", blob[record:record + 12])
        if nid not in (1, 16):
            continue
        raw = blob[offset + strings + off: offset + strings + off + length]
        try:
            best.setdefault(nid, raw.decode("utf-16-be") if pid == 3 else raw.decode("latin-1"))
        except UnicodeDecodeError:
            continue
    return best.get(16) or best.get(1)


def installed_families() -> set[str]:
    """.NET's InstalledFontCollection, which is what GDI can resolve by name."""
    import subprocess
    out = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         "Add-Type -AssemblyName System.Drawing; "
         "(New-Object System.Drawing.Text.InstalledFontCollection).Families "
         "| ForEach-Object { $_.Name }"],
        capture_output=True, text=True, errors="replace",
    ).stdout
    return {line.strip() for line in out.splitlines() if line.strip()}


def cloud_families() -> set[str]:
    """CLOUD DIRECTORIES DO NOT NAME THE FAMILY -- read the name table.

    `CloudFonts\\IBM Plex Sans\\` holds BOTH `IBM Plex Sans` and `IBM Plex Sans
    Condensed`; the directory is the download package, not the family.
    """
    found: set[str] = set()
    if not CLOUD_ROOT.is_dir():
        return found
    for path in CLOUD_ROOT.rglob("*"):
        if path.suffix.lower() in (".ttf", ".otf", ".ttc"):
            name = font_family(path)
            if name:
                found.add(name)
    return found


def requested(deck: Path) -> Counter:
    """Families named on the deck's own SLIDES (theme references excluded)."""
    counts: Counter = Counter()
    with zipfile.ZipFile(deck) as zf:
        for name in zf.namelist():
            if re.match(r"ppt/slides/slide\d+\.xml$", name):
                xml = zf.read(name).decode("utf-8", "replace")
                for match in re.finditer(r'typeface="([^"]+)"', xml):
                    face = match.group(1)
                    if not face.startswith("+"):
                        counts[face] += 1
    return counts


def pdf_families(pdf: Path) -> set[tuple[str, str]]:
    """The families a PDF actually draws with, by POSTSCRIPT NAME.

    ★TRAP: `span["font"]` and `page.get_fonts()`'s BaseFont are NOT the family
    for a PowerPoint export. PowerPoint subsets the fonts it embeds and names
    them by an INDEX -- d16's Source Sans Pro comes back as `41` / `42,Bold`,
    d24's Fira Sans as `57` / `65`. Reading those names made 7 decks look as if
    PowerPoint had fallen back when it had not. Only the embedded file's own
    name table (name ID 6, the PostScript name) identifies the face, so the
    font has to be extracted and parsed.
    """
    seen: set[tuple[str, str]] = set()
    doc = pymupdf.open(pdf)
    try:
        done: set[int] = set()
        for page in doc:
            for font in page.get_fonts(full=False):
                xref, base = font[0], font[3]
                if xref in done:
                    continue
                done.add(xref)
                try:
                    _n, _e, _t, buf = doc.extract_font(xref)
                except (RuntimeError, ValueError):
                    buf = None
                name = postscript_name(buf) or re.sub(r"^[A-Z]{6}\+", "", base)
                seen.add((norm(name), norm(name.split("-")[0])))
    finally:
        doc.close()
    return seen


def postscript_name(blob: bytes) -> str | None:
    """name ID 6 out of an sfnt blob."""
    if blob[:4] not in (SFNT_TAGS):
        return None
    count = struct.unpack(">H", blob[4:6])[0]
    offset = None
    for index in range(count):
        record = 12 + 16 * index
        if blob[record:record + 4] == b"name":
            offset = struct.unpack(">I", blob[record + 8:record + 12])[0]
            break
    if offset is None:
        return None
    _fmt, entries, strings = struct.unpack(">HHH", blob[offset:offset + 6])
    for index in range(entries):
        record = offset + 6 + 12 * index
        pid, _eid, _lid, nid, length, off = struct.unpack(">HHHHHH", blob[record:record + 12])
        if nid != 6:
            continue
        raw = blob[offset + strings + off: offset + strings + off + length]
        try:
            return raw.decode("utf-16-be") if pid == 3 else raw.decode("latin-1")
        except UnicodeDecodeError:
            continue
    return None


def audit() -> list[tuple[str, str, int]]:
    cloud = cloud_families()
    only = {f for f in cloud if f not in installed_families()}
    print(f"cloud-only families ({len(only)}): {', '.join(sorted(only))}\n")
    suspects: list[tuple[str, str, int]] = []
    for deck in sorted(PPTX_DIR.glob("*.pptx")):
        pdf = PDF_DIR / f"{deck.stem}.pdf"
        if not pdf.is_file():
            continue
        asks = {f: c for f, c in requested(deck).items() if norm(f) in {norm(o) for o in only}}
        if not asks:
            continue
        drawn = pdf_families(pdf)
        for face, count in sorted(asks.items(), key=lambda kv: -kv[1]):
            # A PostScript name carries the STYLE too (`OpenSans-Regular`,
            # `FiraSans-Medium`), so the family is a PREFIX of it, never equal —
            # and it ABBREVIATES (`IBMPlexSansCond-Regular` for `IBM Plex Sans
            # Condensed`), so the test has to run both ways: the drawn name may
            # start with the asked-for family, or the family may start with the
            # drawn name's own family part.
            want = norm(face)
            if not any(full.startswith(want) or want.startswith(stem)
                       for full, stem in drawn):
                suspects.append((deck.stem, face, count))
                print(f"  STALE  {deck.stem[:44]:46s} {face:26s} {count:5d} refs")
    print(f"\n{len(suspects)} stale (deck, family) pairs over "
          f"{len({d for d, _f, _c in suspects})} decks")
    return suspects


def refresh(deck_ids: list[str], stamp: str) -> None:
    import win32com.client
    targets: list[Path] = []
    for want in deck_ids:
        hits = sorted(PPTX_DIR.glob(f"{want}__*.pptx")) or sorted(PPTX_DIR.glob(f"{want}.pptx"))
        if not hits:
            sys.exit(f"no deck matched {want}")
        targets.append(hits[0])
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for deck in targets:
            pdf = PDF_DIR / f"{deck.stem}.pdf"
            if pdf.is_file():
                shutil.copy2(pdf, pdf.with_suffix(f".pdf.{stamp}.bak"))
            pres = app.Presentations.Open(str(deck.resolve()), WithWindow=False)
            try:
                pres.SaveAs(str(pdf), 32)  # 32 = ppSaveAsPDF
            finally:
                pres.Close()
            cache = REF_ROOT / deck.stem
            if cache.is_dir():
                shutil.rmtree(cache)
            print(f"  refreshed {deck.stem[:50]}  (reference rasters cleared)")
    finally:
        app.Quit()


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("mode", choices=["audit", "refresh"])
    ap.add_argument("--decks", default=None, help="comma-separated deck ids")
    ap.add_argument("--all-suspect", action="store_true",
                    help="refresh every deck the audit flags")
    ap.add_argument("--stamp", default="bak", help="suffix for the backed-up PDF")
    args = ap.parse_args()
    if args.mode == "audit":
        audit()
        return
    if args.all_suspect:
        ids = sorted({d.split("__")[0] for d, _f, _c in audit()})
    elif args.decks:
        ids = [s.strip() for s in args.decks.split(",") if s.strip()]
    else:
        sys.exit("refresh needs --decks or --all-suspect")
    print(f"\nrefreshing {len(ids)} decks: {', '.join(ids)}")
    refresh(ids, args.stamp)


if __name__ == "__main__":
    main()
