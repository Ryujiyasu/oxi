"""Inspect Meiryo's hwid lookup table to see if it includes yakumono
punctuation chars (、 。「 」 etc.).

Hypothesis: Word's yakumono pair compression uses the font's hwid
(halfwidth) glyph substitution. If Meiryo has halfwidth glyphs for
"、" and "「", Word can compress. If MS Mincho lacks them, Word can't.
"""
from fontTools.ttLib import TTCollection

MEIRYO_TTC = r"C:\Windows\Fonts\meiryo.ttc"

ttc = TTCollection(MEIRYO_TTC)
font = ttc.fonts[0]  # Meiryo regular
gsub = font['GSUB'].table
cmap = font.getBestCmap()

# Find hwid feature
hwid_lookup_indices = []
for fr in gsub.FeatureList.FeatureRecord:
    if fr.FeatureTag == 'hwid':
        hwid_lookup_indices = fr.Feature.LookupListIndex
        break

print(f"Meiryo hwid lookup indices: {hwid_lookup_indices}")

# Yakumono codepoints to test
test_chars = ['、', '。', '，', '．', '「', '」', '『', '』', '（', '）']
test_glyphs = {}
for ch in test_chars:
    cp = ord(ch)
    gid = cmap.get(cp)
    if gid:
        test_glyphs[ch] = gid
print(f"\nTest chars → glyph names: {test_glyphs}")

# Inspect each hwid lookup
for li in hwid_lookup_indices:
    lookup = gsub.LookupList.Lookup[li]
    print(f"\n  Lookup[{li}] type={lookup.LookupType} (1=Single Substitution)")
    for subtable in lookup.SubTable:
        if hasattr(subtable, 'mapping'):
            mapping = subtable.mapping
            print(f"    {len(mapping)} substitutions")
            for ch, gname in test_glyphs.items():
                if gname in mapping:
                    print(f"      '{ch}' ({gname}) → {mapping[gname]}")

# Compare: also check the 'palt' (proportional alternate widths) if Meiryo had it
print(f"\nMeiryo features summary (relevant):")
features_set = set()
for fr in gsub.FeatureList.FeatureRecord:
    features_set.add(fr.FeatureTag)
relevant = ['palt', 'pwid', 'hwid', 'fwid', 'twid', 'qwid', 'halt', 'vhal', 'vpal', 'chws']
for f in relevant:
    has = f in features_set
    print(f"  {f}: {'YES' if has else 'no'}")
