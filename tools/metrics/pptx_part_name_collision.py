# -*- coding: utf-8 -*-
"""Would PowerPoint's own `TTLoadEmbeddedFont` refuse this deck's font parts?

`pptx_embedded_font_name_collision` records that t2embed returns 0x10f when a
part's family name is already taken on the machine, and that the embedded font
then silently becomes the local copy. This asks that question per part, for a
whole deck, WITHOUT rendering anything -- and it asks it under the part's OWN
name (`szWinFamilyName = NULL`), which is what PowerPoint does. Oxi renames every
part to the `p:font/@typeface` the deck's runs ask for, so Oxi's own collisions
are a different question and cannot answer this one.

The name a part carries internally is NOT the one its EOT header declares (the
header describes the SLOT), and it is not readable without decompressing
MicroType Express -- but the refusal itself is the whole answer, so no
decompression is needed.

The Office cloud cache is registered FIRST (`--cloud`, the default), because that
is the state PowerPoint's process is in: a family the cache holds is already
resolvable when the deck's part asks for its name.

    python tools/metrics/pptx_part_name_collision.py 28,18,04,24,15

Verdict per part:
    TAKEN  -> PowerPoint drew the LOCAL copy (system or cloud cache)
    free   -> PowerPoint drew the deck's own part
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
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

TTLOAD_PRIVATE = 0x0000_0001
LICENSE_INSTALLABLE = 0
E_NAME_ALREADY_EXISTS = 0x0000_010F
FR_PRIVATE = 0x10

t2 = ctypes.WinDLL("t2embed.dll")
gdi = ctypes.WinDLL("gdi32.dll")

READEMBEDPROC = ctypes.WINFUNCTYPE(
    ctypes.c_ulong, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_ulong
)


def reader_for(blob: bytes):
    """t2embed pulls the part through a callback; `state` carries the cursor."""
    state = {"pos": 0}

    def read(_stream, buf, count):
        chunk = blob[state["pos"]:state["pos"] + count]
        state["pos"] += len(chunk)
        ctypes.memmove(buf, chunk, len(chunk))
        return len(chunk)

    return READEMBEDPROC(read)


def probe(blob: bytes) -> int:
    handle = wt.HANDLE()
    priv = ctypes.c_ulong()
    status = ctypes.c_ulong()
    cb = reader_for(blob)
    rc = t2.TTLoadEmbeddedFont(
        ctypes.byref(handle), TTLOAD_PRIVATE, ctypes.byref(priv),
        LICENSE_INSTALLABLE, ctypes.byref(status), cb, None,
        None,   # szWinFamilyName -- the part keeps its OWN name
        None, None,
    )
    if rc == 0:
        st = ctypes.c_ulong()
        t2.TTDeleteEmbeddedFont(handle, 0, ctypes.byref(st))
    return rc


def register_cloud(asof: float | None = None) -> list[str]:
    """`asof` registers only the files that existed then, which reconstructs the
    machine PowerPoint saw when it exported a given truth PDF."""
    added = []
    for path in CLOUD.rglob("*"):
        if path.suffix.lower() not in (".ttf", ".otf"):
            continue
        if asof is not None and path.stat().st_mtime > asof:
            continue
        if gdi.AddFontResourceExW(str(path), FR_PRIVATE, None) > 0:
            added.append(str(path))
    return added


def eot_header(blob: bytes) -> tuple[str, str, str]:
    off = 16 + 10 + 2          # sizes/version/flags, PANOSE, charset+italic
    off += 4 + 2 + 2           # weight, fsType, magic
    off += 16 + 8 + 4 + 16     # unicode+codepage ranges, checksum, reserved
    out = []
    for _ in range(3):
        off += 2
        n = struct.unpack_from("<H", blob, off)[0]; off += 2
        out.append(blob[off:off + n].decode("utf-16-le", "replace")); off += n
    return tuple(out)  # FamilyName, StyleName, VersionName


def verdicts(src: Path) -> dict[str, str]:
    out = {}
    with zipfile.ZipFile(src) as z:
        for n in sorted(x for x in z.namelist() if "fntdata" in x):
            rc = probe(z.read(n))
            out[n.split("/")[-1]] = "TAKEN" if rc == E_NAME_ALREADY_EXISTS else (
                "free" if rc == 0 else "rc=0x%x" % rc)
    return out


def main() -> None:
    import datetime as dt
    argv = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    asof = next((a.split("=", 1)[1] for a in flags if a.startswith("--asof=")), None)
    docs = [d.strip() for d in (argv[0] if argv else "").split(",") if d.strip()]
    if not docs:
        sys.exit(__doc__)
    if "--no-cloud" not in flags:
        cut = dt.datetime.strptime(asof, "%Y-%m-%d").timestamp() if asof else None
        n = len(register_cloud(cut))
        if "--json" not in flags:
            print("registered %d cloud files%s" % (n, (" as of " + asof) if asof else ""))
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    out = {}
    for doc in docs:
        src = next((ROOT / "pptx" / i["local"] for i in manifest if i["idx"] == int(doc)), None)
        if src is None or not src.exists():
            print("%s: missing" % doc, file=sys.stderr)
            continue
        v = verdicts(src)
        out["%02d" % int(doc)] = v
        if "--json" in flags:
            continue
        print("=== blind %02d ===" % int(doc))
        with zipfile.ZipFile(src) as z:
            for part, verdict in v.items():
                fam, _sty, ver = eot_header(z.read("ppt/fonts/" + part))
                print("   %-38s declared=%-24r %-22s %s" % (part, fam, ver, verdict))
    if "--json" in flags:
        print(json.dumps(out))


if __name__ == "__main__":
    main()
