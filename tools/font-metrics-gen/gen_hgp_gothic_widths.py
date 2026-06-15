# -*- coding: utf-8 -*-
"""Generate the HGPｺﾞｼｯｸM (HGPGothicM) GDI width table and merge it into
crates/oxidocs-core/src/font/data/gdi_width_overrides.json.

HGPｺﾞｼｯｸM is HG *Proportional* Gothic M — kana/ASCII/punct are proportional
(narrower than fullwidth), kanji are mostly fullwidth. GDI resolves the real
font ONLY via the halfwidth-katakana facename "HGPｺﾞｼｯｸM" (the fullwidth name
"HGPゴシックM" falls back to MS PGothic). Stored under the canonical key
"HGPGothicM" (already in is_cjk_font_family). VERIFIED vs Word PDF render-truth
on kojin: GDI widths = Word natural advances (kana 9.0, kanji 10.5, 、6.75).

Char set = (kojin's unique chars) ∪ (complete proportional ranges) so the
proportional chars are general and kojin's kanji are exact. Missing glyphs are
skipped (GetGlyphIndices) → they fall through to the UPM=256 fullwidth path.
ppem 7-50 (matches the other CJK GDI tables; ppem = round(pt*96/72)).
"""
import ctypes, json, os, sys, zipfile, re, html
sys.stdout.reconfigure(encoding='utf-8')
gdi32 = ctypes.windll.gdi32
FACE = 'HGPｺﾞｼｯｸM'           # the ONLY name GDI resolves to the real font
KEY = 'HGPGothicM'           # canonical table key (is_cjk_font_family)
REPO = r'c:\Users\ryuji\oxi-main'
JSON = os.path.join(REPO, 'crates/oxidocs-core/src/font/data/gdi_width_overrides.json')
DOCX = os.path.join(REPO, 'tools/golden-test/documents/docx/kojin_000505813.docx')
GGI_MARK_NONEXISTING = 1

def kojin_chars():
    z = zipfile.ZipFile(DOCX)
    doc = z.read('word/document.xml').decode('utf-8', 'ignore')
    doc = re.sub(r'<w:instrText[^>]*>.*?</w:instrText>', '', doc, flags=re.S)
    txt = html.unescape(re.sub(r'<[^>]+>', '', doc))
    return {ord(c) for c in txt if ord(c) >= 0x20 and c not in '\t\n\r'}

def charset():
    cps = set(kojin_chars())
    cps |= set(range(0x20, 0x7F))           # ASCII
    cps |= set(range(0x3041, 0x30FF + 1))   # hiragana + katakana
    cps |= set(range(0x3000, 0x303F + 1))   # CJK symbols & punctuation
    cps |= set(range(0xFF01, 0xFFA0))       # fullwidth + halfwidth forms
    cps |= set(range(0xA0, 0x100))          # Latin-1 supplement
    return sorted(cps)

def has_glyph(hdc, cp):
    s = chr(cp)
    idx = (ctypes.c_uint16 * 1)()
    n = gdi32.GetGlyphIndicesW(hdc, s, 1, idx, GGI_MARK_NONEXISTING)
    return n != 0xFFFFFFFF and idx[0] != 0xFFFF

def measure_ppem(cps, ppem):
    hdc = gdi32.CreateCompatibleDC(0)
    hf = gdi32.CreateFontW(-ppem, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, FACE)
    old = gdi32.SelectObject(hdc, hf)
    # confirm the real font resolved
    buf = ctypes.create_unicode_buffer(64); gdi32.GetTextFaceW(hdc, 64, buf)
    assert buf.value == FACE, f"GDI resolved {buf.value!r}, not {FACE!r} — font missing?"
    out = {}
    for cp in cps:
        if not has_glyph(hdc, cp):
            continue
        w = ctypes.c_int(0)
        if gdi32.GetCharWidth32W(hdc, cp, cp, ctypes.byref(w)) and w.value > 0:
            out[str(cp)] = w.value
    gdi32.SelectObject(hdc, old); gdi32.DeleteObject(hf); gdi32.DeleteDC(hdc)
    return out

def main():
    cps = charset()
    print(f"char set: {len(cps)} codepoints; measuring ppem 7-50 for {FACE!r}")
    table = {}
    for ppem in range(7, 51):
        m = measure_ppem(cps, ppem)
        table[str(ppem)] = m
    print(f"ppem14 nchars={len(table['14'])}  sample: A={table['14'].get('65')} "
          f"あ={table['14'].get('12354')} 本={table['14'].get('26412')} 、={table['14'].get('12289')}")
    j = json.load(open(JSON, encoding='utf-8'))
    j[KEY] = table
    json.dump(j, open(JSON, 'w', encoding='utf-8'), ensure_ascii=False)
    print(f"merged {KEY!r} into {JSON} (now {os.path.getsize(JSON)//1024} KB)")

if __name__ == '__main__':
    main()
