# -*- coding: utf-8 -*-
"""Load a deck's embedded parts the way Oxi does, and ask GDI what it got.

`TTLoadEmbeddedFont` returning 0 is not proof the part is in use: GDI can hand
back a substitute under the name it was asked to register. This loads each
`.fntdata` part privately under its `p:font/@typeface` name, then asks the DC
for the face's REAL identity -- `GetTextFace`, the `name` table's full name via
`GetFontData`, and the sfnt magic -- and measures a string, so a face that is
secretly Calibri says so in both the name and the width.

    python tools/metrics/pptx_embedded_face_probe.py 31 --text "Click on the "

Prints, per typeface: the requested name, what GDI reports, the outline format,
and the advance of the probe string at 20.04pt beside the same string in Calibri.
"""
from __future__ import annotations

import argparse
import ctypes as C
import json
import struct
import sys
import zipfile
from ctypes import wintypes as W
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "pipeline_data" / "pptx_benchmark"

t2 = C.WinDLL("t2embed")
gdi = C.WinDLL("gdi32")
READEMBED = C.WINFUNCTYPE(C.c_ulong, C.c_void_p, C.c_void_p, C.c_ulong)


class Stream:
    def __init__(self, data: bytes):
        self.data, self.pos = data, 0

    def read(self, buf, n):
        chunk = self.data[self.pos:self.pos + n]
        C.memmove(buf, chunk, len(chunk))
        self.pos += len(chunk)
        return len(chunk)


def load_part(data: bytes, name: str) -> bool:
    st = Stream(data)
    cb = READEMBED(lambda s, b, n: st.read(b, n))
    h = W.HANDLE()
    priv = C.c_ulong()
    status = C.c_ulong()
    rc = t2.TTLoadEmbeddedFont(C.byref(h), 1, C.byref(priv), 0, C.byref(status),
                               cb, None, C.c_wchar_p(name), None, None)
    load_part.keep = getattr(load_part, "keep", [])
    load_part.keep.append(cb)
    return rc == 0


def face_report(name: str, text: str, size_pt: float) -> dict:
    hdc = gdi.CreateCompatibleDC(None)
    h = gdi.CreateFontW(-int(round(size_pt * 96 / 72 * 10)), 0, 0, 0, 400, 0, 0, 0,
                        1, 0, 0, 4, 0, C.c_wchar_p(name))
    old = gdi.SelectObject(hdc, h)
    got = C.create_unicode_buffer(64)
    gdi.GetTextFaceW(hdc, 64, got)
    magic = (C.c_char * 4)()
    n = gdi.GetFontData(hdc, 0, 0, magic, 4)
    fmt = magic.raw if n == 4 else b"?"
    size = W.SIZE()
    gdi.GetTextExtentPoint32W(hdc, C.c_wchar_p(text), len(text), C.byref(size))
    full = ""
    tag = struct.unpack(">I", b"name")[0]
    ln = gdi.GetFontData(hdc, C.c_uint(0x656D616E), 0, None, 0)  # 'name' little-endian tag
    if ln not in (0, 0xFFFFFFFF):
        buf = (C.c_char * ln)()
        gdi.GetFontData(hdc, C.c_uint(0x656D616E), 0, buf, ln)
        blob = buf.raw
        cnt, so = struct.unpack_from(">HH", blob, 2)
        for i in range(cnt):
            pid, eid, lid, nid, l, off = struct.unpack_from(">6H", blob, 6 + 12 * i)
            if nid == 4:
                s = blob[so + off:so + off + l]
                full = s.decode("utf-16-be", "replace") if pid == 3 else s.decode("latin-1", "replace")
                break
    gdi.SelectObject(hdc, old)
    gdi.DeleteObject(h)
    gdi.DeleteDC(hdc)
    return {"asked": name, "gdi": got.value, "full": full,
            "fmt": fmt.decode("latin-1", "replace"),
            "adv": size.cx / 10 * 72 / 96}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("doc")
    ap.add_argument("--text", default="You need to sign in to your ")
    ap.add_argument("--size", type=float, default=20.04)
    args = ap.parse_args()

    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    src = next(ROOT / "pptx" / i["local"] for i in manifest if f"{i['idx']:02d}" == f"{int(args.doc):02d}")
    z = zipfile.ZipFile(src)
    pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
    rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8", "replace")
    import re
    rid_to_part = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    print(f"probe: {args.text!r} at {args.size}pt")
    print(f"{'typeface':30s} {'GDI face':22s} {'real full name':32s} {'fmt':5s} {'adv':>8s}")
    for m in re.finditer(r'<p:embeddedFont>(.*?)</p:embeddedFont>', pres, re.S):
        blk = m.group(1)
        tf = re.search(r'typeface="([^"]+)"', blk)
        reg = re.search(r'<p:regular r:id="([^"]+)"', blk)
        if not tf or not reg:
            continue
        target = rid_to_part.get(reg.group(1), "")
        part = "ppt/" + target.replace("../", "")
        try:
            data = z.read(part)
        except KeyError:
            continue
        ok = load_part(data, tf.group(1))
        r = face_report(tf.group(1), args.text, args.size)
        print(f"{r['asked']:30s} {r['gdi']:22s} {r['full'][:32]:32s} {r['fmt']:5s} {r['adv']:8.2f}"
              + ("" if ok else "   (TTLoad FAILED)"))
    for ctl in ("Calibri", "Arial", "Segoe UI"):
        r = face_report(ctl, args.text, args.size)
        print(f"{'[control] ' + ctl:30s} {r['gdi']:22s} {r['full'][:32]:32s} {r['fmt']:5s} {r['adv']:8.2f}")


if __name__ == "__main__":
    main()
