# -*- coding: utf-8 -*-
"""Design advances of a privately loaded face, read out of its own tables.

GDI's `GetTextExtentPoint32W` answers with HINTED advances, which differ per
face by up to ~1.5%; the question "which metric did PowerPoint use" cannot be
settled against a ruler that is itself font-dependent. `GetFontData` hands back
the face's `head` / `hhea` / `hmtx` / `cmap` verbatim, so this measures the
design advance the font actually declares.

    python tools/metrics/pptx_hmtx_probe.py 29 --size 14.97 --text "TITLES:"

Loads every embedded part of the deck the way Oxi does, then prints, per face,
the string's design width beside GDI's hinted width at the same size.
"""
from __future__ import annotations

import argparse
import ctypes as C
import json
import re
import struct
import sys
import zipfile
from ctypes import wintypes as W
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx_embedded_face_probe import load_part  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"
gdi = C.WinDLL("gdi32")
PROBE_EM = 2048


def tag(s: str) -> int:
    return struct.unpack("<I", s.encode("latin-1"))[0]


def font_tables(family: str, weight: int) -> dict[str, bytes] | None:
    hdc = gdi.CreateCompatibleDC(None)
    h = gdi.CreateFontW(-PROBE_EM, 0, 0, 0, weight, 0, 0, 0, 1, 0, 0, 4, 0, C.c_wchar_p(family))
    old = gdi.SelectObject(hdc, h)
    out = {}
    for name in ("head", "hhea", "hmtx", "cmap"):
        n = gdi.GetFontData(hdc, C.c_uint(tag(name)), 0, None, 0)
        if n in (0, 0xFFFFFFFF):
            out = {}
            break
        buf = (C.c_char * n)()
        gdi.GetFontData(hdc, C.c_uint(tag(name)), 0, buf, n)
        out[name] = buf.raw
    gdi.SelectObject(hdc, old)
    gdi.DeleteObject(h)
    gdi.DeleteDC(hdc)
    return out or None


def cmap_lookup(cmap: bytes) -> dict[int, int]:
    n = struct.unpack_from(">H", cmap, 2)[0]
    best = None
    for i in range(n):
        pid, eid, off = struct.unpack_from(">HHI", cmap, 4 + 8 * i)
        if (pid, eid) in ((3, 1), (3, 10), (0, 3), (0, 4)):
            best = off
            break
    if best is None:
        return {}
    fmt = struct.unpack_from(">H", cmap, best)[0]
    out: dict[int, int] = {}
    if fmt == 4:
        segx2 = struct.unpack_from(">H", cmap, best + 6)[0]
        seg = segx2 // 2
        ends = struct.unpack_from(f">{seg}H", cmap, best + 14)
        starts = struct.unpack_from(f">{seg}H", cmap, best + 16 + segx2)
        deltas = struct.unpack_from(f">{seg}h", cmap, best + 16 + 2 * segx2)
        range_off_at = best + 16 + 3 * segx2
        offs = struct.unpack_from(f">{seg}H", cmap, range_off_at)
        for i in range(seg):
            for cp in range(starts[i], min(ends[i], 0xFFFF) + 1):
                if offs[i] == 0:
                    g = (cp + deltas[i]) & 0xFFFF
                else:
                    addr = range_off_at + 2 * i + offs[i] + 2 * (cp - starts[i])
                    if addr + 2 > len(cmap):
                        continue
                    g = struct.unpack_from(">H", cmap, addr)[0]
                    if g:
                        g = (g + deltas[i]) & 0xFFFF
                if g:
                    out[cp] = g
    return out


def design_width(family: str, weight: int, text: str, size: float) -> float | None:
    t = font_tables(family, weight)
    if not t:
        return None
    upem = struct.unpack_from(">H", t["head"], 18)[0]
    n_h = struct.unpack_from(">H", t["hhea"], 34)[0]
    cm = cmap_lookup(t["cmap"])
    total = 0
    for ch in text:
        g = cm.get(ord(ch))
        if g is None:
            return None
        i = min(g, n_h - 1)
        if 4 * i + 2 > len(t["hmtx"]):
            return None
        total += struct.unpack_from(">H", t["hmtx"], 4 * i)[0]
    return total * size / upem


def hinted_width(family: str, weight: int, text: str, size: float) -> float:
    hdc = gdi.CreateCompatibleDC(None)
    h = gdi.CreateFontW(-2000, 0, 0, 0, weight, 0, 0, 0, 1, 0, 0, 4, 0, C.c_wchar_p(family))
    old = gdi.SelectObject(hdc, h)
    sz = W.SIZE()
    gdi.GetTextExtentPoint32W(hdc, C.c_wchar_p(text), len(text), C.byref(sz))
    gdi.SelectObject(hdc, old)
    gdi.DeleteObject(h)
    gdi.DeleteDC(hdc)
    return sz.cx / 2000 * size


def load_deck(doc: str) -> list[str]:
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    src = ROOT / "pptx" / next(i["local"] for i in manifest if f"{i['idx']:02d}" == f"{int(doc):02d}")
    z = zipfile.ZipFile(src)
    pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
    rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8", "replace")
    rid = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    out = []
    for m in re.finditer(r"<p:embeddedFont>(.*?)</p:embeddedFont>", pres, re.S):
        blk = m.group(1)
        tf = re.search(r'typeface="([^"]+)"', blk).group(1)
        for slot, suffix in (("regular", ""), ("bold", " Bold"), ("italic", " Italics")):
            r = re.search(rf'<p:{slot} r:id="([^"]+)"', blk)
            if not r:
                continue
            try:
                data = z.read("ppt/" + rid[r.group(1)].replace("../", ""))
            except KeyError:
                continue
            if load_part(data, tf + suffix):
                out.append(tf + suffix)
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("doc")
    ap.add_argument("--text", required=True)
    ap.add_argument("--size", type=float, required=True)
    ap.add_argument("--faces", default="")
    args = ap.parse_args()
    loaded = load_deck(args.doc)
    faces = [f.strip() for f in args.faces.split(",") if f.strip()] or loaded
    print(f"{'face':26s} {'w':>4} {'design':>9} {'hinted':>9}")
    for f in faces:
        for w in (400, 700):
            d = design_width(f, w, args.text, args.size)
            h = hinted_width(f, w, args.text, args.size)
            print(f"  {f:24s} {w:4d} {'--' if d is None else f'{d:9.2f}'} {h:9.2f}")


if __name__ == "__main__":
    main()
