# -*- coding: utf-8 -*-
"""Which embedded parts carry CFF outlines, and which decks draw them?

blind 31 s23 sets a 24pt paragraph in `typeface="Open Sauce"`, the deck embeds a
loadable Open Sauce, and PowerPoint's PDF has the paragraph in **Calibri
6.27** -- the system fallback. Every other family in that deck IS taken from its
embedded part. The one thing that separates them is the outline format: Open
Sauce decompresses to `OTTO` (CFF), the rest to `00 01 00 00` (TrueType).

    PowerPoint does not use an embedded part whose outlines are CFF. It falls
    back as though the family were missing.

★FALSIFIED 2026-09-02. That conclusion was read off the PDF's TEXT LAYER, which
is not its ink. Page 17 of the same deck draws `/Image414 Do` across the
heading's exact box and only then writes `(STRENGTHS)Tj` in F4: a stencil for
the eye, a fallback font for copy-and-paste, because a CFF part is what
PowerPoint cannot EMBED in a PDF -- not what it cannot draw
(`pptx_pdf_stencil_layer`). Rasterise the truth beside a pre-change Oxi render
and both headings are Open Sauce Bold.

Two implementations have now paid for the misreading: dropping the part cost
blind 31 -0.0199 (`_cffskip_ab.log`), and measuring with the part while drawing
Calibri cost it -0.0244, 29 of 33 slides down. What IS true is that the part
decides the BREAK -- Calibri's own advances put 'communication tools' at 206pt
inside boxes of 220 to 243pt, where PowerPoint breaks it anyway -- and that is
what S-FDBREAK reads.

So this census still answers "which parts are CFF"; it does not answer "which
face the deck was drawn in".

The magic number is INSIDE the MicroType Express payload, so it cannot be read
from the EOT header; each part has to be loaded to be classified. Each is deleted
again straight after, so two parts sharing a name cannot shadow each other.

    python tools/metrics/pptx_cff_part_census.py
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import json
import re
import sys
import zipfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"

t2 = ctypes.WinDLL("t2embed.dll")
gdi = ctypes.WinDLL("gdi32.dll")
READ = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_ulong)


def load_under_own_name(blob: bytes):
    state = {"pos": 0}

    def read(_stream, buf, count):
        chunk = blob[state["pos"]:state["pos"] + count]
        state["pos"] += len(chunk)
        ctypes.memmove(buf, chunk, len(chunk))
        return len(chunk)

    cb = READ(read)
    handle = wt.HANDLE()
    priv = ctypes.c_ulong()
    status = ctypes.c_ulong()
    rc = t2.TTLoadEmbeddedFont(ctypes.byref(handle), 1, ctypes.byref(priv), 0,
                               ctypes.byref(status), cb, None, None, None, None)
    return rc, handle, cb


def magic_of(face: str) -> bytes:
    dc = gdi.CreateCompatibleDC(None)
    font = gdi.CreateFontW(-64, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, face)
    old = gdi.SelectObject(dc, font)
    hdr = ctypes.create_string_buffer(4)
    n = gdi.GetFontData(dc, 0, 0, hdr, 4)
    gdi.SelectObject(dc, old)
    gdi.DeleteObject(font)
    gdi.DeleteDC(dc)
    return hdr.raw[:4] if n and n != 0xFFFFFFFF else b""


def declared(blob: bytes) -> str:
    import struct
    off = 16 + 10 + 2 + 4 + 2 + 2 + 16 + 8 + 4 + 16
    off += 2
    n = struct.unpack_from("<H", blob, off)[0]
    return blob[off + 2:off + 2 + n].decode("utf-16-le", "replace")


def main() -> None:
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    hit_decks, total_parts, cff_parts = [], 0, 0
    for item in manifest:
        doc = f"{item['idx']:02d}"
        src = ROOT / "pptx" / item["local"]
        if not src.exists() or not (ROOT / "ssim_pptx" / "ppt_pdf" / f"{doc}.pdf").exists():
            continue
        with zipfile.ZipFile(src) as z:
            drawn = {}
            for n in z.namelist():
                if n.startswith("ppt/slides/slide") and n.endswith(".xml"):
                    for f in re.findall(r'<a:latin typeface="([^"]+)"',
                                        z.read(n).decode("utf-8", "replace")):
                        drawn[f] = drawn.get(f, 0) + 1
            cff = {}
            for n in sorted(x for x in z.namelist() if "fntdata" in x):
                blob = z.read(n)
                total_parts += 1
                face = declared(blob)
                rc, handle, _cb = load_under_own_name(blob)
                if rc != 0:
                    continue
                m = magic_of(face)
                st = ctypes.c_ulong()
                t2.TTDeleteEmbeddedFont(handle, 0, ctypes.byref(st))
                if m == b"OTTO":
                    cff_parts += 1
                    cff[face] = cff.get(face, 0) + 1
        if not cff:
            continue
        used = {f: drawn[f] for f in cff if f in drawn}
        hit_decks.append((doc, cff, used))
        runs = sum(used.values())
        print(f"{doc}  CFF families: {', '.join(sorted(cff))}"
              + (f"   DRAWN in {runs} runs: {used}" if used else "   (never drawn)"))
    print(f"\n{cff_parts} of {total_parts} parts are CFF, in {len(hit_decks)} decks")
    live = [d for d, _c, u in hit_decks if u]
    print(f"{len(live)} decks actually draw one: {','.join(live)}")


if __name__ == "__main__":
    main()
