# -*- coding: utf-8 -*-
"""Export the ascent-split probe and test each candidate metric against it."""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import struct
import sys
from pathlib import Path

import pymupdf
import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = Path(r"pipeline_data\pptx_probes\ascentsplit\ascentsplit.pptx").resolve()
DST = SRC.with_suffix(".pdf")
FONTS = ["Arial", "Goudy Stout", "Castellar", "Stencil", "Maiandra GD",
         "Lucida Handwriting", "Haettenschweiler",
         "Cambria Math", "Noto Serif", "Noto Sans", "Reem Kufi",
         "Liberation Sans Narrow"]
PAIRS = [(20, 60), (60, 20), (20, 20)]

gdi = ctypes.WinDLL("gdi32.dll")
user = ctypes.WinDLL("user32.dll")
gdi.CreateFontW.restype = ctypes.c_void_p
gdi.GetFontData.restype = wt.DWORD


def font_tables(fam: str) -> dict | None:
    hdc = ctypes.c_void_p(user.GetDC(None))
    hf = ctypes.c_void_p(gdi.CreateFontW(-2048, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, fam))
    old = gdi.SelectObject(hdc, hf)
    raw = {}
    try:
        for name in (b"head", b"OS/2", b"hhea"):
            t = struct.unpack("<I", name)[0]
            n = gdi.GetFontData(hdc, t, 0, None, 0)
            if n in (0, 0xFFFFFFFF):
                return None
            buf = ctypes.create_string_buffer(n)
            if gdi.GetFontData(hdc, t, 0, buf, n) == 0xFFFFFFFF:
                return None
            raw[name] = buf.raw
    finally:
        gdi.SelectObject(hdc, old)
        gdi.DeleteObject(hf)
        user.ReleaseDC(None, hdc)
    u16 = lambda b, o: struct.unpack(">H", b[o:o + 2])[0]
    i16 = lambda b, o: struct.unpack(">h", b[o:o + 2])[0]
    head, os2, hhea = raw[b"head"], raw[b"OS/2"], raw[b"hhea"]
    return dict(
        upem=u16(head, 18), ymax=i16(head, 42), ymin=i16(head, 38),
        win_a=u16(os2, 74), win_d=u16(os2, 76),
        typo_a=i16(os2, 68), typo_d=-i16(os2, 70), typo_g=i16(os2, 72),
        hhea_a=i16(hhea, 4), hhea_d=-i16(hhea, 6), hhea_g=i16(hhea, 8),
        fs_sel=u16(os2, 62),
    )


def candidates(m: dict) -> dict[str, float]:
    def r(a, d):
        return 1.2 * a / (a + d) if a + d else float("nan")
    return {
        "win": r(m["win_a"], m["win_d"]),
        "typo": r(m["typo_a"], m["typo_d"]),
        "typo+gap": r(m["typo_a"] + m["typo_g"] / 2, m["typo_d"] + m["typo_g"] / 2),
        "hhea": r(m["hhea_a"], m["hhea_d"]),
        "hhea+gap": r(m["hhea_a"] + m["hhea_g"] / 2, m["hhea_d"] + m["hhea_g"] / 2),
        "bbox": r(m["ymax"], -m["ymin"]),
    }


def export() -> None:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        prs = app.Presentations.Open(str(SRC), WithWindow=False)
        try:
            prs.SaveAs(str(DST), 32)
        finally:
            prs.Close()
    finally:
        app.Quit()


def baselines(page) -> dict[str, float]:
    ys: dict[str, float] = {}
    for blk in page.get_text("rawdict")["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln["spans"]:
                t = "".join(c["c"] for c in sp["chars"]).strip()
                if t in ("AAA", "BBB") and t not in ys:
                    ys[t] = sp["chars"][0]["origin"][1]
    return ys


def main() -> None:
    if "--noexport" not in sys.argv:
        export()
    doc = pymupdf.open(DST)
    keys = ["win", "typo", "typo+gap", "hhea", "hhea+gap", "bbox"]
    header = (f"{'font':22s} {'typo?':>5s} {'lineht':>7s} {'a(meas)':>8s} "
              + " ".join(f"{k:>9s}" for k in keys))
    print(header)
    err = {k: [] for k in keys}
    i = 0
    for font in FONTS:
        steps = {}
        for s1, s2 in PAIRS:
            ys = baselines(doc[i]); i += 1
            steps[(s1, s2)] = ys["BBB"] - ys["AAA"]
        lineht = steps[(20, 20)] / 20.0
        # two mixed arms; each gives a independently, average them
        a_up = (steps[(20, 60)] - lineht * 20) / 40.0
        a_dn = (lineht * 60 - steps[(60, 20)]) / 40.0
        a = (a_up + a_dn) / 2.0
        m = font_tables(font)
        cand = candidates(m) if m else {}
        cells = []
        for k in keys:
            v = cand.get(k, float("nan"))
            cells.append(f"{v:9.4f}")
            if v == v:
                err[k].append(abs(v - a))
        bit = "SET" if m and (m["fs_sel"] & 0x80) else "-"
        print(f"{font[:22]:22s} {bit:>5s} {lineht:7.4f} {a:8.4f} " + " ".join(cells))
    print("\nmean |error| per candidate:")
    for k in keys:
        vals = err[k]
        if vals:
            print(f"  {k:10s} {sum(vals) / len(vals):.4f}   max {max(vals):.4f}")


if __name__ == "__main__":
    main()
