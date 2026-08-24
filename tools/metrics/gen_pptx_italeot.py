# -*- coding: utf-8 -*-
"""Author the EOT-HEADER variants that ask WHY PowerPoint declines an embedded
italic part.

`italface` established the fact: with all four parts embedded, PowerPoint uses
Source Sans Pro's italic parts for every italic request and Barlow's for none.
Both decks' parts are fsType=0 (Installable), flags 0x4, magic 0x504C, with
proper `family` / `style` / `full` strings. Exactly two header fields differ:

    field                 Source Sans Pro          Barlow
    EOT Italic byte       upright 0 -> italic 1    upright 0 -> italic 255
    PANOSE                020b0503030403020204     00000500000000000000
                          -> 020b0503030403090204  -> 00000500000000000000
                             (Letterform 2 -> 9)      (IDENTICAL, all "any")

So Barlow's italic part is indistinguishable from its regular one by PANOSE, and
its Italic byte is out of spec. The EOT header is NOT compressed -- only the
font data after it is -- so both fields can be patched in place without touching
the MicroType Express payload, and PowerPoint asked again.

Arms (each a copy of the Barlow probe with its two italic parts patched):

    eotital   Italic byte 255 -> 1
    eotpan    PANOSE Letterform (byte 7) 0 -> 9
    eotboth   both

If PowerPoint starts using `Barlow-Italic` under one of these, that field is the
gate. If none of them changes anything, the gate is elsewhere and this rules two
candidates out, which is worth as much.

Usage:
    python tools/metrics/gen_pptx_italeot.py
    python tools/metrics/export_pptx_italeot.py
    python tools/metrics/read_pptx_italeot.py
"""
from __future__ import annotations

import json
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
SRC = REPO / "pipeline_data" / "pptx_probes" / "italface" / "italface_Barlow.pptx"
OUT = REPO / "pipeline_data" / "pptx_probes" / "italeot"

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ITALIC_OFF = 27           # EOT header: Italic (1 byte)
PANOSE_OFF = 16           # EOT header: FontPANOSE (10 bytes)
LETTERFORM = 7            # PANOSE[7] = Letterform; 9.. = oblique

ARMS = {
    "eotital": dict(fix_italic=True, fix_panose=False),
    "eotpan": dict(fix_italic=False, fix_panose=True),
    "eotboth": dict(fix_italic=True, fix_panose=True),
}


def italic_part_names(z: zipfile.ZipFile, family: str) -> list[str]:
    rels = {r.get("Id"): r.get("Target")
            for r in etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    out = []
    for f in pres.iter(f"{{{P}}}embeddedFont"):
        fo = f.find(f"{{{P}}}font")
        if fo.get("typeface") != family:
            continue
        for c in f:
            kind = etree.QName(c).localname
            if kind not in ("italic", "boldItalic"):
                continue
            tgt = rels[c.get(f"{{{R}}}id")].lstrip("./")
            out.append(tgt if tgt.startswith("ppt/") else "ppt/" + tgt)
    return out


def patch(data: bytes, fix_italic: bool, fix_panose: bool) -> bytes:
    b = bytearray(data)
    if fix_italic and b[ITALIC_OFF] not in (0, 1):
        b[ITALIC_OFF] = 1
    if fix_panose:
        b[PANOSE_OFF + LETTERFORM] = 9
    return bytes(b)


def main() -> None:
    if not SRC.exists():
        sys.exit(f"missing {SRC} -- run gen_pptx_italface.py first")
    OUT.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(SRC) as z:
        targets = set(italic_part_names(z, "Barlow"))
    print(f"Barlow italic parts: {sorted(targets)}")
    made = []
    for arm, opt in ARMS.items():
        dst = OUT / f"italeot_{arm}.pptx"
        with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in targets:
                    before = data[ITALIC_OFF], data[PANOSE_OFF + LETTERFORM]
                    data = patch(data, **opt)
                    after = data[ITALIC_OFF], data[PANOSE_OFF + LETTERFORM]
                    print(f"  {arm}: {item.filename} italic/letterform {before} -> {after}")
                zout.writestr(item, data)
        made.append(dst.name)
        print(f"wrote {dst}")
    # The arms live on the same slides as the italface Barlow probe.
    idx = json.loads((SRC.parent / "arms.json").read_text(encoding="utf-8"))
    barlow = [e for e in idx if e["family"] == "Barlow"][0]
    (OUT / "arms.json").write_text(
        json.dumps({"files": made, "arms": barlow["arms"]}, indent=1), encoding="utf-8")
    print(f"wrote {OUT / 'arms.json'}")


if __name__ == "__main__":
    main()
