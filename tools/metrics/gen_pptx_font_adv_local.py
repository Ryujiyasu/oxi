# -*- coding: utf-8 -*-
"""Emit design-advance tables for the families the corpora ask for.

The browser has no font system, so `layout::TableMetrics` can only answer for
faces that were measured from their real files and compiled in. Until now that
was three faces, so almost every deck fell back to the browser's own wrap. This
widens the table to the families the pptx corpora actually name, taking each
from the file installed on this machine (system fonts, then the Office cloud
cache).

★The table it writes is for the BROWSER only. The renderer must not consult it:
its advance chain puts `font_adv` ahead of the deck's own embedded part, so a
family added here would SHADOW the part a deck carries -- and those disagree
often (d15 files a Barlow REGULAR in "Barlow Light"'s bold slot, d24 a weight
500 under "Fira Sans Light"). The renderer keeps its richer chain; this is what
remains when there is nothing but the file.

The lookup key is the name a pptx asks for, which is the GDI legacy family:
name ID 1 on its own for the four RIBBI styles, and ID 1 + ID 2 for anything
else -- so Barlow's Light face is asked for as "Barlow Light".

    python tools/metrics/gen_pptx_font_adv_local.py > crates/oxislides-core/src/font_adv_local.rs
"""
from __future__ import annotations

import glob
import os
import re
import sys
import zipfile
from collections import Counter

from fontTools.ttLib import TTFont

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
CORPORA = [
    os.path.join(REPO, "pipeline_data", "pptx_benchmark", "dev", "pptx", "*.pptx"),
    os.path.join(REPO, "pipeline_data", "pptx_benchmark", "pptx", "*.pptx"),
]
ROOTS = [
    r"C:\Windows\Fonts",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "FontCache", "4", "CloudFonts"),
]
RIBBI = {"regular", "bold", "italic", "bold italic"}


def wanted_families() -> Counter:
    """Every `a:latin` typeface the corpora name, by how often."""
    seen: Counter = Counter()
    for pattern in CORPORA:
        for path in glob.glob(pattern):
            try:
                z = zipfile.ZipFile(path)
            except Exception:
                continue
            for name in z.namelist():
                if not name.endswith(".xml"):
                    continue
                if not name.startswith(("ppt/slides/slide", "ppt/slideLayouts/",
                                        "ppt/slideMasters/", "ppt/theme/")):
                    continue
                xml = z.read(name).decode("utf-8", "replace")
                for m in re.finditer(r'<a:latin typeface="([^"]+)"', xml):
                    face = m.group(1)
                    if not face.startswith("+"):
                        seen[face] += 1
    return seen


def local_faces() -> dict:
    """(asked-for family, bold, italic) -> font path, for every installed face."""
    out: dict = {}
    for root in ROOTS:
        if not root or not os.path.isdir(root):
            continue
        for path in glob.glob(os.path.join(root, "**", "*.*"), recursive=True):
            if os.path.splitext(path)[1].lower() not in (".ttf", ".otf", ".ttc"):
                continue
            try:
                font = TTFont(path, lazy=True, fontNumber=0)
                names = font["name"]
                family = names.getDebugName(1) or ""
                sub = (names.getDebugName(2) or "").strip()
                mac_style = font["head"].macStyle
            except Exception:
                continue
            if not family:
                continue
            # ★The style comes from the BITS, not from the subfamily string.
            # A file's name can lie about what it is: this machine's
            # `Caladea-BoldItalic.ttf` calls itself just "Italic" (name ID 2)
            # while its `macStyle` correctly says bold+italic, and reading the
            # string filed the bold-italic face as the deck's plain italic.
            # Same shape as the embedded parts whose declared typeface is not
            # what they hold -- the name is not the identity.
            bold = bool(mac_style & 0x01)
            italic = bool(mac_style & 0x02)
            # The key is the name a pptx asks for: the GDI legacy family, which
            # is ID 1 alone for the four RIBBI styles and ID 1 + ID 2 for the
            # rest -- so Barlow's Light face is asked for as "Barlow Light".
            low = sub.lower()
            asked = family if low in RIBBI else f"{family} {sub}".strip()
            out.setdefault((asked, bold, italic), path)
    return out


def table_for(path: str) -> list | None:
    """ASCII 32..126 advances in EM, or None if the face cannot serve them."""
    try:
        font = TTFont(path, lazy=True, fontNumber=0)
        upm = font["head"].unitsPerEm
        cmap = font.getBestCmap()
        hmtx = font["hmtx"]
    except Exception:
        return None
    row = []
    for cp in range(32, 127):
        glyph = cmap.get(cp)
        if glyph is None or glyph not in hmtx.metrics:
            return None
        row.append(round(hmtx[glyph][0] / upm, 5))
    return row


def main() -> None:
    asked = wanted_families()
    faces = local_faces()
    rows = []
    for family, _uses in asked.most_common():
        for bold in (False, True):
            for italic in (False, True):
                path = faces.get((family, bold, italic))
                if path is None:
                    continue
                row = table_for(path)
                if row is None:
                    continue
                rows.append((family, bold, italic, row, os.path.basename(path)))

    ident = {}
    print("// This Source Code Form is subject to the terms of the Mozilla Public")
    print("// License, v. 2.0. If a copy of the MPL was not distributed with this")
    print("// file, You can obtain one at https://mozilla.org/MPL/2.0/.")
    print()
    print("//! Design advances for the families the pptx corpora name, measured from")
    print("//! the files installed on the machine that generated this.")
    print("//!")
    print("//! GENERATED by `tools/metrics/gen_pptx_font_adv_local.py` -- do not edit.")
    print("//!")
    print("//! ★For the BROWSER only. The renderer's advance chain puts the measured")
    print("//! tables AHEAD of a deck's own embedded part, so consulting this there")
    print("//! would shadow the part the deck carries -- and those disagree often")
    print("//! (d15 files a Barlow Regular in \"Barlow Light\"'s bold slot). This is")
    print("//! what a build with no font system has instead, not a better answer.")
    print("//!")
    print("//! Each table is ASCII 32..=126 in code-point order, in EM units.")
    print()
    for family, bold, italic, row, src in rows:
        const = re.sub(r"[^A-Z0-9]", "_", family.upper())
        const += ("_B" if bold else "") + ("_I" if italic else "")
        while const in ident:
            const += "_"
        ident[const] = True
        print(f"/// {family}{' Bold' if bold else ''}{' Italic' if italic else ''} ({src}).")
        print(f"static {const}: [f32; 95] = [")
        for i in range(0, 95, 10):
            chunk = ", ".join(f"{v:.5f}" for v in row[i:i + 10])
            print(f"    {chunk},")
        print("];")
        print()
    print("/// The measured face for one request, or None when nothing was measured.")
    print("fn table(family: &str, bold: bool, italic: bool) -> Option<&'static [f32; 95]> {")
    print("    let key = family.to_ascii_lowercase();")
    print("    Some(match (key.as_str(), bold, italic) {")
    for (family, bold, italic, _row, _src), const in zip(rows, ident.keys()):
        print(f'        ("{family.lower()}", {str(bold).lower()}, {str(italic).lower()}) => &{const},')
    print("        _ => return None,")
    print("    })")
    print("}")
    print()
    print("/// The design advance of `ch` for this face, in EM units.")
    print("pub fn local_advance_em(family: &str, bold: bool, italic: bool, ch: char) -> Option<f32> {")
    print("    let idx = ch as u32 as usize;")
    print("    if !(32..127).contains(&idx) {")
    print("        return None;")
    print("    }")
    print("    // A style the machine did not have falls back to the upright face,")
    print("    // which is what a browser drawing a synthesised bold also does.")
    print("    let t = table(family, bold, italic)")
    print("        .or_else(|| table(family, bold, false))")
    print("        .or_else(|| table(family, false, false))?;")
    print("    Some(t[idx - 32])")
    print("}")
    print()
    print("/// Whether any face of `family` was measured.")
    print("pub fn local_family_supported(family: &str) -> bool {")
    print("    table(family, false, false).is_some()")
    print("        || table(family, true, false).is_some()")
    print("}")
    print(f"\n// {len(rows)} faces measured, from {len({r[0] for r in rows})} families.", file=sys.stderr)


if __name__ == "__main__":
    main()
