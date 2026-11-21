"""Build crates/oxidocs-core/src/font/data/vertical_font_metrics.json from the installed fonts.

A vertical column advances each glyph by its vmtx advance AFTER the GSUB 'vert'/'vrt2'
substitution (small kana, punctuation and the long-vowel mark take their vertical forms).
Word does this for proportional faces: ＭＳ Ｐ明朝 and HGP創英角ﾎﾟｯﾌﾟ体 18pt in a tbRl column
advance ぐ 17.28 (vmtx 246/256), っ 9.72 (vert form 138/256), し 16.32 (232).

Only faces with proportional kana (some of あいうえおアイウエオ narrower than the em in hmtx)
are written; a monospaced face advances one em and needs no table. The four ＭＳ faces keep
their full tables; the other faces are written sparse (only advances that differ from the
em, `default_advance` = em) -- the metrics only, never the font file.

usage: python tools/metrics/gen_vertical_font_metrics.py [--write]
"""
import glob, json, os, sys
from fontTools.ttLib import TTCollection, TTFont

OUT = os.path.join(os.path.dirname(__file__), "..", "..", "crates", "oxidocs-core", "src", "font", "data",
                   "vertical_font_metrics.json")
FULL = {"MS Mincho", "MS PMincho", "MS Gothic", "MS PGothic"}
PROBE = "あいうえおアイウエオ"


def vert_map(f):
    m = {}
    if "GSUB" not in f:
        return m
    g = f["GSUB"].table
    idx = set()
    for fr in g.FeatureList.FeatureRecord:
        if fr.FeatureTag in ("vert", "vrt2"):
            idx.update(fr.Feature.LookupListIndex)
    for i in sorted(idx):
        lk = g.LookupList.Lookup[i]
        for st in lk.SubTable:
            if lk.LookupType == 7:
                st = st.ExtSubTable
            if hasattr(st, "mapping"):
                m.update(st.mapping)
    return m


def faces():
    for path in sorted(glob.glob("C:/Windows/Fonts/*.tt[cf]") + glob.glob("C:/Windows/Fonts/*.TT[CF]")):
        try:
            fonts = TTCollection(path).fonts if path.lower().endswith(".ttc") else [TTFont(path)]
        except Exception:
            continue
        for f in fonts:
            yield f


def main():
    data = json.load(open(OUT, encoding="utf-8"))
    seen = set(data)
    for f in faces():
        try:
            name = f["name"].getDebugName(1)
        except Exception:
            continue
        if name in seen or "vmtx" not in f:
            continue
        cmap = f.getBestCmap() or {}
        upm = f["head"].unitsPerEm
        hm = f["hmtx"].metrics
        if not any(ord(c) in cmap and hm[cmap[ord(c)]][0] < 0.98 * upm for c in PROBE):
            continue
        vm = f["vmtx"].metrics
        vmap = vert_map(f)
        adv = {}
        for cp, gl in cmap.items():
            if cp < 0x2000:
                continue
            a = vm.get(vmap.get(gl, gl), vm.get(gl, (upm, 0)))[0]
            if a != upm:
                adv[cp] = a
        data[name] = {"units_per_em": upm, "default_advance": upm, "advances": adv}
        seen.add(name)
        print(name, upm, len(adv), "ぐ", adv.get(ord("ぐ"), upm), "っ", adv.get(ord("っ"), upm), file=sys.stderr)
    if "--write" in sys.argv:
        with open(OUT, "w", encoding="utf-8", newline="\n") as fh:
            json.dump(data, fh, ensure_ascii=False, separators=(",", ":"))


if __name__ == "__main__":
    main()
