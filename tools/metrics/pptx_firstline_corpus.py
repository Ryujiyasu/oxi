# -*- coding: utf-8 -*-
"""Score the first-baseline rule against the whole corpus, and say which set of
vertical metrics PowerPoint read.

The `firstline` probe derived the rule on INSTALLED faces and concluded the
source is `usWin`. Two corpus decks disagree, both in embedded **Bebas Neue**,
whose `usWin` gives 0.8769 and whose `sTypo` gives 0.9000 -- and 0.9000 is also
the value at which the rule collapses to the pre-change `0.75 * P * n`, so one
font cannot tell those two readings apart. The corpus can: several embedded
faces have `fsSelection` bit 7 CLEAR (so the renderer reads `usWin`) while their
`sTypo` says something quite different.

For every text shape that is simple enough to measure without a layout engine --
`anchor="t"`, a single paragraph, one `lnSpc`, text that appears verbatim in
PowerPoint's own PDF -- this reports

    measured_off = pdf_baseline - (box_top + tIns)

against the rule evaluated with each metric source, and tallies which source is
closer. No Oxi render is involved, so nothing here can be contaminated by the
renderer's own state.

Usage:
    python tools/metrics/pptx_firstline_corpus.py
    python tools/metrics/pptx_firstline_corpus.py --decks d20,d38 --verbose
"""
from __future__ import annotations

import argparse
import re
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
EMU = 12700.0


def sfnt_tables(blob: bytes) -> dict[bytes, tuple[int, int]]:
    count = struct.unpack(">H", blob[4:6])[0]
    tabs: dict[bytes, tuple[int, int]] = {}
    for index in range(count):
        rec = 12 + 16 * index
        tabs[blob[rec:rec + 4]] = struct.unpack(">II", blob[rec + 8:rec + 16])
    return tabs


def vertical_metrics(blob: bytes) -> dict | None:
    """PostScript name plus the three candidate ascent splits."""
    if blob[:4] not in (bytes([0, 1, 0, 0]), b"OTTO", b"true"):
        return None
    tabs = sfnt_tables(blob)
    if b"OS/2" not in tabs or b"hhea" not in tabs:
        return None
    os2 = tabs[b"OS/2"][0]
    hhea = tabs[b"hhea"][0]
    u16 = lambda k: struct.unpack(">H", blob[os2 + k:os2 + k + 2])[0]
    i16 = lambda k: struct.unpack(">h", blob[os2 + k:os2 + k + 2])[0]
    hi = lambda k: struct.unpack(">h", blob[hhea + k:hhea + k + 2])[0]
    typo_asc, typo_desc = i16(68) + i16(72), -i16(70)
    name = "?"
    if b"name" in tabs:
        base = tabs[b"name"][0]
        _fmt, entries, strings = struct.unpack(">HHH", blob[base:base + 6])
        for index in range(entries):
            rec = base + 6 + 12 * index
            pid, _eid, _lid, nid, length, off = struct.unpack(">HHHHHH", blob[rec:rec + 12])
            if nid != 6:
                continue
            raw = blob[base + strings + off: base + strings + off + length]
            try:
                name = (raw.decode("utf-16-be") if pid == 3 else raw.decode("latin-1"))
                name = name.replace("\x00", "")
            except UnicodeDecodeError:
                pass
            break
    return {
        "name": name,
        "fs_selection": u16(62),
        "win": 1.2 * u16(74) / (u16(74) + u16(76)) if u16(74) + u16(76) else None,
        "typo": 1.2 * typo_asc / (typo_asc + typo_desc) if typo_asc + typo_desc > 0 else None,
        "hhea": 1.2 * hi(4) / (hi(4) - hi(6)) if hi(4) - hi(6) else None,
    }


def rule(face: float, size: float, n: float) -> float:
    pitch = 1.2 * size
    natural = pitch - face * size
    quarter = 0.25 * pitch
    if n <= 1.0:
        descent = max(natural + quarter * (n - 1.0), min(natural, quarter * n))
    else:
        descent = max(natural, quarter * n)
    return pitch * n - descent


def shapes_of(xml: str) -> list[dict]:
    """The text shapes simple enough to measure: anchor=t, one paragraph."""
    out: list[dict] = []
    for match in re.finditer(r"<p:sp>.*?</p:sp>", xml, re.S):
        body = re.sub(r"\s+", " ", match.group(0))
        if body.count("<a:p>") != 1:
            continue
        body_pr = re.search(r"<a:bodyPr[^>]*>", body)
        if not body_pr or 'anchor="t"' not in body_pr.group(0):
            continue
        if "<a:normAutofit" in body:
            continue
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"\s*/>', body)
        if not off:
            continue
        t_ins = re.search(r'tIns="(-?\d+)"', body_pr.group(0))
        pct = re.findall(r'<a:lnSpc><a:spcPct val="(\d+)"/></a:lnSpc>', body)
        size = re.findall(r'sz="(\d+)"', body)
        face = re.findall(r'<a:latin typeface="([^"]+)"', body)
        text = "".join(re.findall(r"<a:t>([^<]*)</a:t>", body))
        if not text.strip() or not size or not face:
            continue
        out.append({
            "top": int(off.group(2)) / EMU + (int(t_ins.group(1)) / EMU if t_ins else 3.6),
            "n": int(pct[0]) / 100000.0 if pct else 1.0,
            "size": int(size[0]) / 100.0,
            "face": face[0],
            "text": text.strip(),
            "explicit_t_ins": bool(t_ins),
        })
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--decks", default=None)
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()
    wanted = {s.strip() for s in args.decks.split(",")} if args.decks else None

    tally: Counter = Counter()
    per_face: dict[str, Counter] = {}
    for deck in sorted(DEV.joinpath("pptx").glob("*.pptx")):
        if wanted and deck.stem.split("__")[0] not in wanted:
            continue
        pdf = DEV / "pdf" / f"{deck.stem}.pdf"
        if not pdf.is_file():
            continue
        zf = zipfile.ZipFile(deck)
        doc = pymupdf.open(pdf)
        try:
            for page_index in range(doc.page_count):
                name = f"ppt/slides/slide{page_index + 1}.xml"
                if name not in zf.namelist():
                    continue
                page = doc[page_index]
                # PostScript name per PDF font resource, so a span can be tied
                # to the metrics PowerPoint actually embedded.
                by_ref: dict[str, dict] = {}
                for font in page.get_fonts(full=False):
                    try:
                        _n, _e, _t, blob = doc.extract_font(font[0])
                    except (RuntimeError, ValueError):
                        continue
                    if not blob:
                        continue
                    met = vertical_metrics(blob)
                    if met:
                        by_ref[font[4]] = met
                spans = []
                for block in page.get_text("rawdict")["blocks"]:
                    for line in block.get("lines", []):
                        for span in line["spans"]:
                            spans.append((
                                "".join(c["c"] for c in span["chars"]).strip(),
                                span["origin"][1],
                                span["size"],
                                by_ref.get(span.get("xref_name", ""), None),
                                span["font"],
                            ))
                xml = zf.read(name).decode("utf-8", "replace")
                for shape in shapes_of(xml):
                    hit = None
                    for text, y, size, met, raw in spans:
                        if not text or not shape["text"].startswith(text[:12]):
                            continue
                        if abs(size - shape["size"]) > 0.6:
                            continue
                        hit = (y, size, raw)
                        break
                    if hit is None:
                        continue
                    measured = hit[0] - shape["top"]
                    # Find the embedded font whose PostScript name matches the
                    # requested family, ignoring spaces and the style suffix.
                    key = re.sub(r"[^a-z0-9]", "", shape["face"].lower())
                    met = None
                    for candidate in by_ref.values():
                        stem = re.sub(r"[^a-z0-9]", "", candidate["name"].split("-")[0].lower())
                        if stem and (stem.startswith(key) or key.startswith(stem)):
                            met = candidate
                            break
                    if not met or met["win"] is None or met["typo"] is None:
                        continue
                    errs = {
                        src: abs(measured - rule(met[src], hit[1], shape["n"]))
                        for src in ("win", "typo", "hhea")
                        if met[src] is not None
                    }
                    if min(errs.values()) > 3.0 and not args.verbose:
                        continue  # the shape is not the one the span belongs to
                    best = min(errs, key=lambda k: errs[k])
                    tally[best] += 1
                    per_face.setdefault(met["name"], Counter())[best] += 1
                    if args.verbose:
                        print(f"  {deck.stem[:10]:12s} s{page_index + 1:<3d} "
                              f"{met['name'][:22]:24s} fsSel=0x{met['fs_selection']:04x} "
                              f"sz={hit[1]:6.2f} n={shape['n']:.4f} meas={measured:8.3f} "
                              + "  ".join(f"{k}={errs[k]:6.3f}" for k in sorted(errs))
                              + f"  -> {best}")
        finally:
            doc.close()
    print("\nclosest metric source, over", sum(tally.values()), "measurable shapes:")
    for src, count in tally.most_common():
        print(f"  {src:5s} {count}")
    print("\nfaces where the verdict is not unanimous:")
    for face, counts in sorted(per_face.items()):
        if len(counts) > 1:
            print(f"  {face:26s} {dict(counts)}")


if __name__ == "__main__":
    main()
