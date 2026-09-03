# -*- coding: utf-8 -*-
"""Rebuild a .pptx whose download was cut before the zip's central directory.

Blind decks 16 and 22 are both **exactly 11,907,091 bytes** and both end without
an EOCD record, so `zipfile` refuses them and the corpus has run with 112 of its
114 decks since. A zip's central directory lives at the END of the file, but
every member is also introduced by its own local header, so what survives the
cut is still readable by walking forward from byte 0.

These members carry no sizes in the local header (general-purpose bit 3: the
sizes follow the data, in a descriptor), so each entry has to be decompressed
until its deflate stream ends -- the decompressor's own `unused_data` is what
says where the member stopped.

Writes `<name>.recovered.pptx` beside the original and reports which parts were
recovered, so a deck missing `ppt/presentation.xml` can be told from one missing
only a trailing image.

    python tools/metrics/pptx_recover_truncated.py <file.pptx> [more...]
    python tools/metrics/pptx_recover_truncated.py --corpus
"""
from __future__ import annotations

import glob
import os
import struct
import sys
import zipfile
import zlib

SIG_LOCAL = b"PK\x03\x04"


def members(data: bytes):
    """Every member the file still holds, as (name, bytes), forward from 0."""
    pos = 0
    n = len(data)
    while pos + 30 <= n:
        if data[pos:pos + 4] != SIG_LOCAL:
            # A cut can leave a partial member; skip to the next header.
            nxt = data.find(SIG_LOCAL, pos + 1)
            if nxt < 0:
                return
            pos = nxt
            continue
        (_, flags, method, _, _, crc, csize, usize, name_len, extra_len) = struct.unpack(
            "<HHHHHIIIHH", data[pos + 4:pos + 30]
        )
        start = pos + 30 + name_len + extra_len
        if start > n:
            return
        name = data[pos + 30:pos + 30 + name_len].decode(
            "utf-8" if flags & 0x800 else "cp437", "replace"
        )
        if method == 0 and csize:
            body = data[start:start + csize]
            pos = start + csize
            if len(body) == csize:
                yield name, body
            continue
        if method != 8:
            return
        d = zlib.decompressobj(-zlib.MAX_WBITS)
        try:
            body = d.decompress(data[start:])
        except zlib.error:
            return
        if not d.eof:
            # The file ends inside this member: nothing after it is recoverable.
            return
        consumed = len(data) - start - len(d.unused_data)
        after = start + consumed
        # An optional data descriptor follows; it may or may not be signed.
        if data[after:after + 4] == b"PK\x07\x08":
            after += 16
        elif csize == 0 and usize == 0:
            after += 12
        pos = after
        if name:
            yield name, body


DEFAULTS = {
    "rels": "application/vnd.openxmlformats-package.relationships+xml",
    "xml": "application/xml",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
    "svg": "image/svg+xml",
    "mp4": "video/mp4",
    "m4a": "audio/mp4",
    "wav": "audio/wav",
    "fntdata": "application/x-fontdata",
    "bin": "application/vnd.openxmlformats-officedocument.oleObject",
    "thmx": "application/vnd.ms-officetheme",
}
PML = "application/vnd.openxmlformats-officedocument.presentationml."
DML = "application/vnd.openxmlformats-officedocument.drawingml."
OVERRIDES = [
    ("ppt/presentation.xml", PML + "presentation.main+xml"),
    ("ppt/slides/", PML + "slide+xml"),
    ("ppt/slideLayouts/", PML + "slideLayout+xml"),
    ("ppt/slideMasters/", PML + "slideMaster+xml"),
    ("ppt/notesSlides/", PML + "notesSlide+xml"),
    ("ppt/notesMasters/", PML + "notesMaster+xml"),
    ("ppt/handoutMasters/", PML + "handoutMaster+xml"),
    ("ppt/commentAuthors.xml", PML + "commentAuthors+xml"),
    ("ppt/presProps.xml", PML + "presProps+xml"),
    ("ppt/viewProps.xml", PML + "viewProps+xml"),
    ("ppt/tableStyles.xml", PML + "tableStyles+xml"),
    ("ppt/theme/", DML + "theme+xml"),
    ("ppt/charts/", DML + "chart+xml"),
    ("ppt/diagrams/", DML + "diagramData+xml"),
    ("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
    ("docProps/app.xml", PML.replace("presentationml.", "") + "extended-properties+xml"),
]


def content_types(names: list[str]) -> bytes:
    """The part-type index, rebuilt from the parts themselves.

    It is a manifest, not content: every entry is determined by the part's own
    path and extension, which is why a cut that takes only this part is
    recoverable at all. Parts whose type cannot be named that way (an unknown
    extension) are left out, and the package still opens as long as nothing
    references them.
    """
    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">']
    exts = sorted({n.rsplit(".", 1)[-1].lower() for n in names if "." in n})
    for ext in exts:
        if ext in DEFAULTS:
            lines.append('<Default Extension="%s" ContentType="%s"/>' % (ext, DEFAULTS[ext]))
    for name in sorted(names):
        if not name.endswith(".xml") or "/_rels/" in name:
            continue
        for prefix, ctype in OVERRIDES:
            if name == prefix or (prefix.endswith("/") and name.startswith(prefix)):
                lines.append('<Override PartName="/%s" ContentType="%s"/>' % (name, ctype))
                break
    lines.append("</Types>")
    return "".join(lines).encode("utf-8")


# A 1x1 fully transparent PNG. Stands in for a picture the cut removed.
BLANK_PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4"
    "890000000a49444154789c6360000002000100ffff03000006000557bfabd400"
    "00000049454e44ae426082"
)


def missing_targets(parts: dict[str, bytes]) -> set[str]:
    """Parts that surviving relationships still name but the cut took.

    Dropping the relationship is not enough: the slide that used it still
    carries `r:embed="rId3"`, and PowerPoint reads a package whose shape names a
    relationship that is not there as corrupt (`0x80070570`). So the
    relationship is kept and its target supplied empty -- both engines then read
    the same file and draw the same nothing, which is what a comparison corpus
    needs. The picture is lost either way; what is preserved is the deck's text
    and geometry.
    """
    import posixpath
    import re

    want = set()
    for name, body in parts.items():
        if not name.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(name))
        text = body.decode("utf-8", "replace")
        for m in re.finditer(r"<Relationship\b[^>]*/>", text):
            if 'TargetMode="External"' in m.group(0):
                continue
            tgt = re.search(r'Target="([^"]+)"', m.group(0))
            if not tgt:
                continue
            t = tgt.group(1)
            path = t[1:] if t.startswith("/") else posixpath.normpath(posixpath.join(base, t))
            if path not in parts:
                want.add(path)
    return want


def drop_dangling(rels: bytes, have: set[str], base: str) -> bytes:
    """Relationships pointing at parts the cut removed, taken out.

    `docProps/app.xml` and `core.xml` sit at the tail of these two files, so the
    package relationships still name them. PowerPoint refuses a package whose
    relationship targets a part that is not there.
    """
    import re

    text = rels.decode("utf-8", "replace")

    def keep(m: "re.Match[str]") -> str:
        target = re.search(r'Target="([^"]+)"', m.group(0))
        mode = re.search(r'TargetMode="([^"]+)"', m.group(0))
        if not target or (mode and mode.group(1) == "External"):
            return m.group(0)
        t = target.group(1)
        if t.startswith("/"):
            path = t[1:]
        else:
            path = os.path.normpath(os.path.join(base, t)).replace(os.sep, "/")
        return m.group(0) if path in have else ""

    return re.sub(r"<Relationship\b[^>]*/>", keep, text).encode("utf-8")


def recover(path: str) -> None:
    data = open(path, "rb").read()
    got = list(members(data))
    out = os.path.splitext(path)[0] + ".recovered.pptx"
    if not got:
        print("%-52s nothing readable" % os.path.basename(path))
        return
    import re

    parts = dict(got)
    filled = sorted(missing_targets(parts))
    # An embedded font is declared as well as related, and PowerPoint reads an
    # empty one as a corrupt package. A picture can stand in blank; a font has
    # to be un-declared, so the deck falls back the way a machine without it
    # would.
    lost_fonts = [f for f in filled if f.lower().endswith(".fntdata")]
    if lost_fonts and "ppt/presentation.xml" in parts:
        rels_name = "ppt/_rels/presentation.xml.rels"
        rels = parts.get(rels_name, b"").decode("utf-8", "replace")
        drop_ids = {
            m.group(1)
            for m in re.finditer(r'<Relationship\b[^>]*Id="([^"]+)"[^>]*/>', rels)
            if (t := re.search(r'Target="([^"]+)"', m.group(0)))
            and "ppt/" + t.group(1).replace("../", "") in lost_fonts
        }
        pres = parts["ppt/presentation.xml"].decode("utf-8", "replace")
        kept = []
        for block in re.finditer(r"<p:embeddedFont>.*?</p:embeddedFont>", pres, re.S):
            if any(rid in block.group(0) for rid in drop_ids):
                kept.append(block.group(0))
        for block in kept:
            pres = pres.replace(block, "")
        pres = re.sub(r"<p:embeddedFontLst>\s*</p:embeddedFontLst>", "", pres)
        parts["ppt/presentation.xml"] = pres.encode("utf-8")
        parts[rels_name] = re.sub(
            r'<Relationship\b[^>]*/>',
            lambda m: "" if re.search(r'Id="([^"]+)"', m.group(0)).group(1) in drop_ids else m.group(0),
            rels,
        ).encode("utf-8")
        filled = [f for f in filled if f not in lost_fonts]
    for name in filled:
        parts[name] = BLANK_PNG if name.lower().endswith((".png", ".jpg", ".jpeg")) else b""
    names = list(parts)
    if "[Content_Types].xml" not in parts:
        parts["[Content_Types].xml"] = content_types(names)
    order = ["[Content_Types].xml"] + [n for n in parts if n != "[Content_Types].xml"]
    have = set(parts)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order:
            body = parts[name]
            if name.endswith(".rels"):
                base = os.path.dirname(os.path.dirname(name))
                body = drop_dangling(body, have, base)
            z.writestr(name, body)
    names = list(parts)
    if filled:
        print("%-52s %d part(s) supplied empty: %s"
              % ("", len(filled), ", ".join(os.path.basename(f) for f in filled[:6])))
    need = ["[Content_Types].xml", "ppt/presentation.xml"]
    missing = [p for p in need if p not in names]
    slides = sum(1 for n in names if n.startswith("ppt/slides/slide"))
    media = sum(1 for n in names if n.startswith("ppt/media/"))
    print("%-52s %4d parts, %2d slides, %3d media%s"
          % (os.path.basename(path), len(names), slides, media,
             "" if not missing else "  MISSING " + ", ".join(missing)))
    if missing:
        os.remove(out)
        return
    try:
        with zipfile.ZipFile(out) as z:
            bad = z.testzip()
        print("%-52s -> %s%s" % ("", os.path.basename(out), "" if bad is None else "  BAD " + bad))
    except Exception as exc:  # noqa: BLE001
        print("%-52s -> rebuilt file will not open: %s" % ("", exc))


def main() -> None:
    args = sys.argv[1:]
    if "--corpus" in args:
        root = os.path.join("pipeline_data", "pptx_benchmark", "pptx")
        args = []
        for p in sorted(glob.glob(os.path.join(root, "*.pptx"))):
            if p.endswith(".recovered.pptx"):
                continue
            try:
                with zipfile.ZipFile(p):
                    continue
            except Exception:  # noqa: BLE001
                args.append(p)
        if not args:
            print("every corpus deck opens as a zip")
            return
    for path in args:
        recover(path)


if __name__ == "__main__":
    main()
