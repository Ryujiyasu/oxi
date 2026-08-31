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


# Beyond ASCII, the characters Western text is actually set with: the Latin-1
# Supplement, and the punctuation and symbols a word processor inserts.
#
# ★Deliberately NOT the set our own corpora happen to use. A table shaped to
# this corpus would raise our own numbers and do nothing for somebody else's
# deck; this range is the one that generalises. What the corpora DO show is
# that the gap is real and small: 82 distinct non-ASCII characters in 1761
# occurrences, led by U+2019 (21 decks), U+2039/U+203A, U+2714, U+00AE, U+2014
# and the accented Latin letters -- every one of them inside this set except
# the emoji, which need a different face anyway.
EXTRA_CPS = (
    list(range(0x00A0, 0x0100))
    + [0x2013, 0x2014, 0x2018, 0x2019, 0x201A, 0x201C, 0x201D, 0x201E,
       0x2020, 0x2021, 0x2022, 0x2026, 0x2030, 0x2039, 0x203A,
       0x20AC, 0x2122, 0x2212, 0x25A0, 0x25AA, 0x25C6, 0x25CF, 0x2713, 0x2714]
)


def table_for(path: str) -> tuple | None:
    """ASCII 32..126 advances in EM plus whatever of EXTRA_CPS the face has.

    None when the face cannot serve the ASCII range -- a face that cannot set
    plain text is no use to anyone. The extras are per-face and may be empty:
    a character the face lacks is left out so the engine declines the run
    instead of advancing it by a glyph that is not there.
    """
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
    extra = []
    for cp in EXTRA_CPS:
        glyph = cmap.get(cp)
        if glyph is None or glyph not in hmtx.metrics:
            continue
        extra.append((cp, round(hmtx[glyph][0] / upm, 5)))
    return row, extra


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
                got = table_for(path)
                if got is None:
                    continue
                row, extra = got
                rows.append((family, bold, italic, row, extra, os.path.basename(path)))

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
    print("//! Each dense table is ASCII 32..=126 in code-point order, in EM units.")
    print("//! A `_X` table beside it carries what that face has beyond ASCII -- the")
    print("//! Latin-1 Supplement and the punctuation Western text is set with --")
    print("//! sparse, sorted by code point, and holding only the characters the face")
    print("//! really contains.")
    print()
    for family, bold, italic, row, extra, src in rows:
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
        if extra:
            print(f"/// {family} beyond ASCII, sorted by code point.")
            print(f"static {const}_X: &[(u32, f32)] = &[")
            for i in range(0, len(extra), 5):
                chunk = ", ".join(f"({cp}, {v:.5f})" for cp, v in extra[i:i + 5])
                print(f"    {chunk},")
            print("];")
            print()
    print("/// The measured face for one request, or None when nothing was measured.")
    print("fn table(family: &str, bold: bool, italic: bool) -> Option<&'static [f32; 95]> {")
    print("    let key = family.to_ascii_lowercase();")
    print("    Some(match (key.as_str(), bold, italic) {")
    for (family, bold, italic, _row, _extra, _src), const in zip(rows, ident.keys()):
        print(f'        ("{family.lower()}", {str(bold).lower()}, {str(italic).lower()}) => &{const},')
    print("        _ => return None,")
    print("    })")
    print("}")
    print()
    print("/// The same face's advances beyond ASCII, or an empty slice.")
    print("fn extras(family: &str, bold: bool, italic: bool) -> &'static [(u32, f32)] {")
    print("    let key = family.to_ascii_lowercase();")
    print("    match (key.as_str(), bold, italic) {")
    for (family, bold, italic, _row, extra, _src), const in zip(rows, ident.keys()):
        if extra:
            print(f'        ("{family.lower()}", {str(bold).lower()}, '
                  f'{str(italic).lower()}) => {const}_X,')
    print("        _ => &[],")
    print("    }")
    print("}")
    print()
    print("/// The design advance of `ch` for this face, in EM units.")
    print("pub fn local_advance_em(family: &str, bold: bool, italic: bool, ch: char) -> Option<f32> {")
    print("    let cp = ch as u32;")
    print("    // ★Only the style that was actually measured answers. Serving a")
    print("    // bold request from the upright face looks harmless and is not:")
    print("    // Merriweather Bold is about 1% wider than its Regular, which put")
    print("    // d08's 38pt titles up to 9pt past where PowerPoint drew them, and")
    print("    // the layout still called itself complete. Six of the seventeen")
    print("    // families here have no bold face on this machine; their bold text")
    print("    // is declined rather than set on the wrong advances. Whether a")
    print("    // SYNTHESISED bold even advances like its upright is unmeasured")
    print("    // (S577c, parked), so there is nothing to fall back ON.")
    print("    if (32..127).contains(&cp) {")
    print("        let t = table(family, bold, italic)?;")
    print("        return Some(t[cp as usize - 32]);")
    print("    }")
    print("    // Beyond ASCII the tables are sparse: a face carries only the")
    print("    // characters it actually has, so a miss here is a real answer --")
    print("    // the caller declines the run rather than advancing a glyph that")
    print("    // is not in the font.")
    print("    let src = extras(family, bold, italic);")
    print("    match src.binary_search_by_key(&cp, |e| e.0) {")
    print("        Ok(i) => Some(src[i].1),")
    print("        Err(_) => None,")
    print("    }")
    print("}")
    print()
    print("/// Whether any face of `family` was measured.")
    print("pub fn local_family_supported(family: &str) -> bool {")
    print("    table(family, false, false).is_some()")
    print("        || table(family, true, false).is_some()")
    print("}")
    print(f"\n// {len(rows)} faces measured, from {len({r[0] for r in rows})} families;"
          f" {sum(len(r[4]) for r in rows)} advances beyond ASCII.", file=sys.stderr)


if __name__ == "__main__":
    main()
