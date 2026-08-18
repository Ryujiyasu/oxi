# -*- coding: utf-8 -*-
"""Export an embedded-font probe and check the ascent split against OS/2.

The prediction under test (probe `ascentsplit`, 12 installed faces):

    ascent = usWinAscent, descent = usWinDescent                      normally
    ascent = sTypoAscender + sTypoLineGap, descent = -sTypoDescender
                          when OS/2 fsSelection bit 7 (USE_TYPO_METRICS) is set
    a = 1.2 * ascent / (ascent + descent)

The OS/2 values are read from the font PowerPoint itself embedded in the
exported PDF, so there is no question about which face got resolved.

    python tools/metrics/read_pptx_embedsplit.py
    python tools/metrics/read_pptx_embedsplit.py --name embedsplit_d15 \
        --fonts "Barlow,Barlow Light,Montserrat"
"""
from __future__ import annotations

import argparse
import struct
import sys
from pathlib import Path

import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PROBES = Path(r"pipeline_data\pptx_probes").resolve()
PAIRS = [(20, 60), (60, 20), (20, 20), (30, 50)]


def export(src: Path, dst: Path) -> None:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        prs = app.Presentations.Open(str(src), WithWindow=False)
        try:
            prs.SaveAs(str(dst), 32)
        finally:
            prs.Close()
    finally:
        app.Quit()


def read_os2(data: bytes) -> dict | None:
    if len(data) < 12:
        return None
    n = struct.unpack(">H", data[4:6])[0]
    tabs = {}
    for i in range(n):
        o = 12 + 16 * i
        if o + 16 > len(data):
            return None
        tabs[data[o:o + 4].decode("latin1")] = struct.unpack(">II", data[o + 8:o + 16])
    if "OS/2" not in tabs or "head" not in tabs:
        return None
    u16 = lambda b, o: struct.unpack(">H", b[o:o + 2])[0]
    i16 = lambda b, o: struct.unpack(">h", b[o:o + 2])[0]
    oo, ho = tabs["OS/2"][0], tabs["head"][0]
    return dict(
        upem=u16(data, ho + 18),
        fs_sel=u16(data, oo + 62),
        win=(u16(data, oo + 74), u16(data, oo + 76)),
        typo=(i16(data, oo + 68), -i16(data, oo + 70), i16(data, oo + 72)),
    )


def predicted(t: dict) -> float:
    if t["fs_sel"] & 0x80:
        a, d = t["typo"][0] + t["typo"][2], t["typo"][1]
    else:
        a, d = t["win"]
    return 1.2 * a / (a + d) if a + d else float("nan")


def face_tables(doc) -> dict[str, dict]:
    """PDF font resource name -> OS/2 of the face PowerPoint embedded for it."""
    out: dict[str, dict] = {}
    for pi in range(doc.page_count):
        for xref, *_ in ((f[0],) for f in doc[pi].get_fonts(full=True)):
            obj = doc.xref_object(xref)
            name = obj.split("/BaseFont /")[1].split()[0] if "/BaseFont /" in obj else str(xref)
            if name in out:
                continue
            for cand in doc.xref_get_keys(xref):
                pass
            # FontFile2 lives on the descriptor, one or two hops down
            import re
            m = re.search(r"/FontDescriptor\s*(\d+) 0 R", obj)
            if not m:
                m2 = re.search(r"/DescendantFonts\s*\[\s*(\d+) 0 R", obj)
                if not m2:
                    continue
                obj2 = doc.xref_object(int(m2.group(1)))
                m = re.search(r"/FontDescriptor\s*(\d+) 0 R", obj2)
                if not m:
                    continue
            desc = doc.xref_object(int(m.group(1)))
            mf = re.search(r"/FontFile2\s*(\d+) 0 R", desc)
            if not mf:
                continue
            t = read_os2(doc.xref_stream(int(mf.group(1))))
            if t:
                out[name] = t
    return out


def spans(page) -> dict[str, tuple[float, str]]:
    out: dict[str, tuple[float, str]] = {}
    for blk in page.get_text("rawdict")["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln["spans"]:
                t = "".join(c["c"] for c in sp["chars"]).strip()
                if t in ("AAA", "BBB") and t not in out:
                    out[t] = (sp["chars"][0]["origin"][1], sp["font"])
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--name", default="embedsplit")
    ap.add_argument("--fonts", default="Calistoga,Jua")
    ap.add_argument("--noexport", action="store_true")
    args = ap.parse_args()
    fonts = [f.strip() for f in args.fonts.split(",")]
    src = PROBES / args.name / f"{args.name}.pptx"
    dst = src.with_suffix(".pdf")
    if not args.noexport:
        export(src, dst)
    doc = pymupdf.open(dst)
    tables = face_tables(doc)
    i = 0
    for font in fonts:
        steps, faces = {}, set()
        for s1, s2 in PAIRS:
            d = spans(doc[i]); i += 1
            steps[(s1, s2)] = d["BBB"][0] - d["AAA"][0]
            faces.add(d["AAA"][1])
        lineht = steps[(20, 20)] / 20.0
        est = [(steps[(s1, s2)] - lineht * s1) / (s2 - s1) for s1, s2 in PAIRS if s1 != s2]
        a = sum(est) / len(est)
        print(f"\n{font}  (PDF faces {sorted(faces)})")
        print(f"  line height {lineht:.4f}   a = {a:.4f}   arms {['%.4f' % e for e in est]}")
        for face in sorted(faces):
            t = next((v for k, v in tables.items() if k.endswith(face) or face in k), None)
            if not t:
                print(f"    {face}: no embedded OS/2 in the PDF")
                continue
            p = predicted(t)
            branch = "typo" if t["fs_sel"] & 0x80 else "win"
            print(f"    {face}: upem={t['upem']} win={t['win']} typo={t['typo']} "
                  f"[{branch}] predicted {p:.4f}   error {abs(p - a):.4f}")


if __name__ == "__main__":
    main()
