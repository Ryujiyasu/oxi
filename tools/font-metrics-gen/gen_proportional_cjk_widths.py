# -*- coding: utf-8 -*-
"""Measure a PROPORTIONAL Japanese face's per-character GDI advances and merge
them into crates/oxidocs-core/src/font/data/gdi_width_overrides.json.gz.

Generalises `gen_hgp_gothic_widths.py`, which did exactly this for HGPｺﾞｼｯｸM.

Why: a proportional Japanese face narrows kana, marks and some ideographs; the
compact metrics table carries only ~100 Latin widths per family, so every CJK
character falls through to the fullwidth em and every line wraps early.

  technical__898a80c889101e85 (BIZ UDPゴシック 18pt, column 425.20pt)
      Word  25 chars/line, advances 13.68 / 16.20 / 16.56 / 16.78 / 18.00
      Oxi   23 chars/line, advances **18.00 x 22** (fullwidth for all of them)
  -> Oxi wraps two characters early on every line, +1 page over Word's 14.

`char_width_pt_with_gdi_map` already consults this table before falling back to
the em, so a family present here needs no code change at all.

The population is not small: of 569 corpus docs, **50** set a proportional
Japanese face as the body's dominant eastAsia font, and only three families
(MS PGothic / MS PMincho / HGPGothicM) are covered today -- 22 docs are not.

Usage:
    gen_proportional_cjk_widths.py "BIZ UDPゴシック" [table-key]
    gen_proportional_cjk_widths.py --list        # families still missing

The GDI face name must be the one GDI actually resolves (GetTextFaceW is
asserted against it -- HGPｺﾞｼｯｸM only resolves under its halfwidth-katakana
spelling, and the fullwidth one silently falls back to MS PGothic).
"""
import ctypes
import gzip
import json
import os
import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
gdi32 = ctypes.windll.gdi32
GZ = _REPO / "crates" / "oxidocs-core" / "src" / "font" / "data" / "gdi_width_overrides.json.gz"
GGI_MARK_NONEXISTING = 1


def charset():
    """Everything a proportional Japanese face narrows, plus the common kanji.

    Kanji are mostly fullwidth even in a proportional face, but 'mostly' is not
    'always' (the technical doc measures 16.56 and 16.78 on ideographs at 18pt),
    so the CJK block is measured rather than assumed. Glyphs the face lacks are
    skipped and keep falling through to the em.
    """
    cps = set(range(0x20, 0x7F))            # ASCII
    cps |= set(range(0xA0, 0x100))          # Latin-1 supplement
    cps |= set(range(0x2000, 0x206F))       # general punctuation
    cps |= set(range(0x3000, 0x303F + 1))   # CJK symbols & punctuation
    cps |= set(range(0x3041, 0x30FF + 1))   # hiragana + katakana
    cps |= set(range(0x4E00, 0x9FA0))       # CJK unified ideographs
    cps |= set(range(0xFF01, 0xFFA0))       # fullwidth + halfwidth forms
    return sorted(cps)


def has_glyph(hdc, cp):
    idx = (ctypes.c_uint16 * 1)()
    n = gdi32.GetGlyphIndicesW(hdc, chr(cp), 1, idx, GGI_MARK_NONEXISTING)
    return n != 0xFFFFFFFF and idx[0] != 0xFFFF


def measure_ppem(face, cps, ppem):
    hdc = gdi32.CreateCompatibleDC(0)
    hf = gdi32.CreateFontW(-ppem, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, face)
    old = gdi32.SelectObject(hdc, hf)
    buf = ctypes.create_unicode_buffer(64)
    gdi32.GetTextFaceW(hdc, 64, buf)
    if buf.value != face:
        gdi32.SelectObject(hdc, old)
        gdi32.DeleteObject(hf)
        gdi32.DeleteDC(hdc)
        raise SystemExit("GDI resolved %r, not %r -- font missing or wrong spelling"
                         % (buf.value, face))
    out = {}
    w = ctypes.c_int(0)
    for cp in cps:
        if not has_glyph(hdc, cp):
            continue
        if gdi32.GetCharWidth32W(hdc, cp, cp, ctypes.byref(w)) and w.value > 0:
            out[str(cp)] = w.value
    gdi32.SelectObject(hdc, old)
    gdi32.DeleteObject(hf)
    gdi32.DeleteDC(hdc)
    return out


def load():
    with gzip.open(GZ, "rt", encoding="utf-8") as f:
        return json.load(f)


def save(j):
    with gzip.open(GZ, "wt", encoding="utf-8", compresslevel=9) as f:
        json.dump(j, f, ensure_ascii=False)


def main():
    if len(sys.argv) < 2 or sys.argv[1] == "--list":
        j = load()
        print("gdi_width_overrides に収録済み: %d family" % len(j))
        for k in sorted(j):
            print("   ", k)
        return
    face = sys.argv[1]
    key = sys.argv[2] if len(sys.argv) > 2 else face
    cps = charset()
    print("face=%r key=%r  char set=%d codepoints; ppem 7-50" % (face, key, len(cps)))
    table = {}
    for ppem in range(7, 51):
        table[str(ppem)] = measure_ppem(face, cps, ppem)
        if ppem == 24:
            m = table["24"]
            print("  ppem24 nchars=%d  A=%s あ=%s 本=%s 、=%s ス=%s"
                  % (len(m), m.get("65"), m.get("12354"), m.get("26412"),
                     m.get("12289"), m.get("12473")))
    j = load()
    j[key] = table
    save(j)
    print("merged %r -> %s (%.1f MB)" % (key, GZ, os.path.getsize(GZ) / 1048576.0))


if __name__ == "__main__":
    main()
