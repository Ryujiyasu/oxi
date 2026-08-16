"""Add real Tw Cen MT metrics to font_metrics_compact.json.

forms__002fbe2c (blindC50) sets `w:ascii="Garamond"` on 1098 runs -- the whole
body -- but the registry had no Garamond entry, so every line fell back to the
Calibri default (natural 1.2207em) where Word uses Garamond (1.125em).  That is
-0.096em per line: a 10pt line 12.207 -> 11.25, i.e. ~-1pt per line, which
accumulates over a 50-line page and spills the last one or two paragraphs.

Same shape as S819 (Segoe UI) / S855 (Arial Narrow) / S860 (Bookman) /
S866 (Aptos) / S950 (Book Antiqua) / S1001 (Open Sans) / S1007 (Comic Sans) /
S1012 (Century Gothic) / S1032 (Bell MT) / S1033 (Palatino) / S1060 (Verdana) /
S1086 (Calibri Light): pure data, no code change -- normalize/render are
passthrough and `registry.get("Garamond")` hits on the exact name.

usage: python tools/metrics/add_twcen_metrics.py [--dry-run]
"""

import json
import os
import sys

JSON = os.path.join(
    "crates", "oxidocs-core", "src", "font", "data", "font_metrics_compact.json"
)
FACES = [
    ("Tw Cen MT", r"C:\Windows\Fonts\TCM_____.TTF"),
    # S1138 (2026-08-15): Trebuchet MS is installed on every Windows box and named
    # by 4 corpus docs, but had no entry, so its lines fell back to Calibri --
    # _pb_linepitch_gen.py measures Word 1.16132em against Oxi's 1.22070, i.e.
    # 0.59pt on a 10pt line. The face's own hhea gives 1.16113, one 600-DPI
    # quantum from Word's figure.
    ("Trebuchet MS", "C:/Windows/Fonts/trebuc.ttf"),
    ("Trebuchet MS Bold", "C:/Windows/Fonts/trebucbd.ttf"),
    # S1139 (2026-08-15): Meiryo UI is a separate face inside meiryo.ttc, not a
    # style of Meiryo -- its descent is 430 against Meiryo's 901, so the natural
    # height is 1.27002em, not 1.5. Without an entry the name normalised to
    # Meiryo and every line came out 1.5 x 83/64 = 1.94531em where Word draws
    # 1.27002 x 83/64 = 1.65103 (_pb_linepitch_gen.py measures 1.65119). 5 docs.
    ("Meiryo UI", "C:/Windows/Fonts/meiryo.ttc#2"),
    # S1140 (2026-08-15): the rest of the sweep -- every family the corpus
    # names in a body/header/footer run that is INSTALLED here but had no
    # entry, so Oxi measured it as Calibri. Found by diffing the corpus's
    # rFonts names against the table and the machine's font directory;
    # each is verified against Word by _pb_linepitch_gen.py after the
    # rebuild. CJK names (Yu Gothic / SimSun / MS UI Gothic) are left out --
    # they already resolve through normalize_family_name.
    ("Segoe UI Emoji", "C:/Windows/Fonts/seguiemj.ttf"),
    ("Ink Free", "C:/Windows/Fonts/Inkfree.ttf"),
    ("Franklin Gothic Book", "C:/Windows/Fonts/FRABK.TTF"),
    ("Wingdings", "C:/Windows/Fonts/wingding.ttf"),
    ("Sylfaen", "C:/Windows/Fonts/sylfaen.ttf"),
    ("Lucida Sans Unicode", "C:/Windows/Fonts/l_10646.ttf"),
    ("Jokerman", "C:/Windows/Fonts/JOKERMAN.TTF"),
    ("Impact", "C:/Windows/Fonts/impact.ttf"),
    ("Eras Bold ITC", "C:/Windows/Fonts/ERASBD.TTF"),
    ("Broadway", "C:/Windows/Fonts/BROADW.TTF"),
    ("Baskerville Old Face", "C:/Windows/Fonts/BASKVILL.TTF"),
    ("Arial Rounded MT Bold", "C:/Windows/Fonts/ARLRDBD.TTF"),
    # S1141 (2026-08-15): the last family in the sweep. Its only face declares
    # subfamily "Italic" (the design is a calligraphic slant), which is why the
    # resolver pass that found the others skipped it -- the file is still what
    # Word uses for `w:ascii="Lucida Calligraphy"`. reference__00476cb8 sets it
    # on 18 runs, and the Calibri fallback it was getting is 1.22070 against the
    # face's 1.36133: +1.55pt on an 11pt line.
    ("Lucida Calligraphy", "C:/Windows/Fonts/LCALLIG.TTF"),
    # S1142 (2026-08-15): the two font sources the earlier audit missed --
    # OpenType (.otf, invisible to a *.tt* glob) and Office's CLOUD font cache.
    # Word renders these as the real face, so the table needs their metrics.
    # Latin cloud faces, each verified against Word in a COMPLETE document
    # (the minimal-docx probe cannot be trusted for font resolution):
    ("Montserrat", "cloud:Montserrat"),                 # Word 1.21900 (typo set)
    ("Merriweather", "cloud:Merriweather"),             # Word 1.25757 (typo set)
    ("Nunito", "cloud:Nunito"),                         # Word 1.36417 (typo set)
    ("Roboto", "cloud:Roboto"),                         # Word 1.20117
    ("Source Sans Pro", "cloud:Source Sans Pro"),       # Word 1.25691
    ("Avenir Next LT Pro", "cloud:Avenir Next LT Pro"), # Word 1.21308
    # UNHELD 2026-08-16 after S1145 gave `line=0 atLeast` its exact natural
    # height inside a typed grid -- the rule that governs every body
    # paragraph these faces move in educational__0214ac95.
    ("PMingLiU", "cloud:PMingLiU", True),
    ("Batang", "cloud:Batang", True),
    ("UD デジタル 教科書体 N-R", "C:/Windows/Fonts/UDDigiKyokashoN-R.ttc#0", True),
    ("UD デジタル 教科書体 NP-R", "C:/Windows/Fonts/UDDigiKyokashoN-R.ttc#1", True),
    ("UD デジタル 教科書体 NK-R", "C:/Windows/Fonts/UDDigiKyokashoN-R.ttc#2", True),
    ("UD デジタル 教科書体 N-B", "C:/Windows/Fonts/UDDigiKyokashoN-B.ttc#0", True),
    ("UD デジタル 教科書体 NP-B", "C:/Windows/Fonts/UDDigiKyokashoN-B.ttc#1", True),
    ("UD デジタル 教科書体 NK-B", "C:/Windows/Fonts/UDDigiKyokashoN-B.ttc#2", True),
    ("BIZ UDゴシック", "C:/Windows/Fonts/BIZ-UDGothicR.ttc#0", True),
    ("BIZ UDPゴシック", "C:/Windows/Fonts/BIZ-UDGothicR.ttc#1", True),
    ("BIZ UD明朝 Medium", "C:/Windows/Fonts/BIZ-UDMinchoM.ttc#0", True),
    # S1147 (2026-08-16): the HG family. The earlier sweep missed these
    # because its scan read only w:ascii / w:hAnsi -- a Japanese document
    # names its body face in w:eastAsia, and HG丸ｺﾞｼｯｸM-PRO alone carries
    # 6744 references across 14 corpus documents. All are installed here, so
    # Word draws them for real; without an entry they took the unresolved
    # path, which S1146 sends to Yu Gothic -- 6 JA blind documents lost a
    # page each until these landed. Widths empty per the S579 CJK convention.
    ("HG丸ｺﾞｼｯｸM-PRO", "C:/Windows/Fonts/HGRSMP.TTF", True),
    ("HGP創英角ｺﾞｼｯｸUB", "C:/Windows/Fonts/HGRSGU.TTC#1", True),
    ("HGPｺﾞｼｯｸE", "C:/Windows/Fonts/HGRGE.TTC#1", True),
    ("HGP行書体", "C:/Windows/Fonts/HGRGY.TTC#1", True),
    ("HGS明朝E", "C:/Windows/Fonts/HGRME.TTC#2", True),
    ("HG創英角ﾎﾟｯﾌﾟ体", "C:/Windows/Fonts/HGRPP1.TTC#0", True),
    ("HGP創英角ﾎﾟｯﾌﾟ体", "C:/Windows/Fonts/HGRPP1.TTC#1", True),
    # S1148 (2026-08-16): MS UI Gothic -- the last hole the audit found. It is
    # installed (msgothic.ttc face 1) and Word draws it at 1.29717em, but it
    # was never in the table and its name carries no CJK characters, so
    # S1146 routed it to Cambria (1.17237) -- the widest miss of the sweep.
    ("MS UI Gothic", "C:/Windows/Fonts/msgothic.ttc#1", True),
]
# the codepoint set every existing entry carries
REF_FAMILY = "Verdana"


def resolve(path):
    r"""Expand a "cloud:<Family>" spec to a real file.

    Office keeps its downloadable faces in
    %LOCALAPPDATA%\Microsoft\FontCache\CloudFonts\<Family>\<numeric id>.ttf --
    a third font source besides C:\Windows\Fonts and the per-user font dir, and
    one that neither a *.tt* glob nor .NET's InstalledFontCollection reports.
    Word renders these as the real face (measured: it embeds "Avenir Next LT Pro"
    itself, not a substitute), so their metrics belong in the table. The numeric
    file name differs per machine, hence the lookup by family directory.
    """
    if not path.startswith("cloud:"):
        return path
    fam = path[len("cloud:"):]
    root = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "FontCache",
                        "4", "CloudFonts", fam)
    if not os.path.isdir(root):
        return path
    from fontTools.ttLib import TTFont
    best = None
    for f in sorted(os.listdir(root)):
        if not f.lower().endswith((".ttf", ".otf")):
            continue
        full = os.path.join(root, f)
        try:
            sub = (TTFont(full, lazy=True)["name"].getDebugName(2) or "").lower()
        except Exception:
            continue
        if sub in ("regular", "book", ""):
            return full
        best = best or full
    return best or path


def extract(path, codepoints):
    from fontTools.ttLib import TTFont

    # "path#N" selects a face inside a TrueType Collection (meiryo.ttc holds
    # Meiryo and Meiryo UI as separate faces, not as styles of one family).
    num = 0
    if "#" in path:
        path, _, n = path.partition("#")
        num = int(n)
    f = TTFont(path, fontNumber=num, lazy=True)
    head, hhea, os2 = f["head"], f["hhea"], f["OS/2"]
    upm = head.unitsPerEm
    cmap = f.getBestCmap()
    if cmap is None:
        # A symbol font (Wingdings) ships only the (3,0) symbol cmap, where the
        # ASCII range lives at 0xF000 + cp. getBestCmap() rejects it outright.
        sym = next((t.cmap for t in f["cmap"].tables
                    if t.platformID == 3 and t.platEncID == 0), None)
        cmap = {cp: sym[0xF000 + cp] for cp in range(0x20, 0x7F)
                if sym and (0xF000 + cp) in sym} if sym else {}
    hmtx = f["hmtx"]
    widths = {}
    missing = []
    for cp in codepoints:
        g = cmap.get(int(cp))
        if g is None:
            missing.append(cp)
            continue
        widths[str(cp)] = int(round(hmtx[g][0] * 2048 / upm))
    e = {
        "family": None,
        "units_per_em": 2048,
        "ascender": int(round(hhea.ascent * 2048 / upm)),
        "descender": int(round(hhea.descent * 2048 / upm)),
        "line_gap": int(round(hhea.lineGap * 2048 / upm)),
        "win_ascent": int(round(os2.usWinAscent * 2048 / upm)),
        "win_descent": int(round(os2.usWinDescent * 2048 / upm)),
        "typo_ascender": int(round(os2.sTypoAscender * 2048 / upm)),
        "typo_descender": int(round(os2.sTypoDescender * 2048 / upm)),
        "typo_line_gap": int(round(os2.sTypoLineGap * 2048 / upm)),
        # S1142 (2026-08-15): fsSelection bit 7 (USE_TYPO_METRICS) decides which
        # set Word measures the line with. Measured on the Office CLOUD fonts,
        # where the two sets differ widely: bit set -> the typo sum (Montserrat
        # 1.21900 = Word 1.21900, Merriweather 1.25700 = 1.25757, Nunito 1.36400
        # = 1.36417); bit clear -> max(hhea+gap, win), the S950 rule (Roboto
        # 1.20020 = 1.20117, Source Sans Pro 1.25700 = 1.25691, Avenir Next LT
        # Pro 1.21289 = 1.21308). Without this field those three would each be
        # 0.6-1.6pt per 10pt line too tall.
        "use_typo_metrics": bool(os2.fsSelection & (1 << 7)),
        "widths": widths,
    }
    f.close()
    return e, missing


def main():
    dry = "--dry-run" in sys.argv
    data = json.load(open(JSON, encoding="utf-8"))
    before = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    ref = [x for x in data if x.get("family") == REF_FAMILY][0]
    codepoints = sorted(int(k) for k in ref["widths"])
    have = {x.get("family") for x in data}

    for face in FACES:
        fam, path = face[0], face[1]
        no_widths = len(face) > 2 and face[2]
        if fam in have:
            print("already present:", fam)
            continue
        path = resolve(path)
        if not os.path.exists(path.partition("#")[0]):
            print("MISSING FONT FILE:", path)
            return 1
        e, missing = extract(path, [] if no_widths else codepoints)
        e["family"] = fam
        nat = (
            (e["typo_ascender"] + abs(e["typo_descender"]) + e["typo_line_gap"])
            if e["use_typo_metrics"]
            else max(
                (e["ascender"] + abs(e["descender"]) + e["line_gap"]),
                (e["win_ascent"] + e["win_descent"]),
            )
        ) / 2048.0
        print(
            "%-16s upm=2048 hhea=%d/%d/%d win=%d/%d natural=%.6f widths=%d missing=%d"
            % (
                fam,
                e["ascender"],
                e["descender"],
                e["line_gap"],
                e["win_ascent"],
                e["win_descent"],
                nat,
                len(e["widths"]),
                len(missing),
            )
        )
        data.append(e)

    if dry:
        print("(dry run, not written)")
        return 0

    # additive only: every pre-existing entry must be byte-identical
    out = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    kept = json.loads(out)[: len(json.loads(before))]
    assert json.dumps(kept, ensure_ascii=False, separators=(",", ":")) == before, (
        "existing entries changed!"
    )
    open(JSON, "w", encoding="utf-8").write(out)
    print("wrote", JSON, "entries:", len(data))
    return 0


if __name__ == "__main__":
    sys.exit(main())
