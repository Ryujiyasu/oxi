"""Extract REAL Aptos / Aptos Display metrics without the font file on disk.

Aptos is the Office 2024+ default theme font (theme1.xml <a:latin
typeface="Aptos"/> via docDefaults asciiTheme="minorHAnsi"). Word ships and
renders it, but it is NOT a loose .ttf anywhere on C: (a system-wide search
finds none) -- so the usual "parse C:/Windows/Fonts/<face>.ttf" route
(S819 Segoe UI / S855 Arial Narrow / S860 Bookman) is unavailable.

Route used here: author a docx that sets rFonts ascii/hAnsi="Aptos" on a run
containing EVERY codepoint we need, export it via Word COM to PDF (Word
embeds a SUBSET containing exactly the glyphs used), then extract that
embedded font and read head/hhea/OS-2 + hmtx with fontTools. Because the
probe text covers the whole target set, the subset's coverage == the target
set. head/hhea/OS-2 are font-wide and survive subsetting intact.

Verified against the reference__0014acda Word PDF: Aptos upm 2048,
hhea 1923/-577/0 (line 1.2207em == Calibri's), win 2068/563, and glyphs
~11% WIDER than Calibri ('a' 0.5312em vs 0.4790em) -- which is why an
Aptos doc falling back to Calibri metrics under-reserves its wrap.

Usage:
  python _font_aptos_extract.py            # gen docx -> Word PDF -> metrics
  python _font_aptos_extract.py --emit     # also print the compact JSON entries
"""
import os, sys, json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_font_aptos")

# Match the codepoint coverage of the existing compact entries (Calibri):
# printable ASCII + smart quotes + circled numbers.
TARGET = ([c for c in range(32, 127)]
          + [8216, 8217, 8220, 8221]
          + list(range(9312, 9332)))

FACES = {
    "Aptos": dict(font="Aptos", bold=False),
    "Aptos Bold": dict(font="Aptos", bold=True),
    "Aptos Display": dict(font="Aptos Display", bold=False),
}


def probe_text():
    # every target codepoint, space-separated so Word cannot ligate/kern them
    # into one another; the space itself (32) is in the set too.
    return " ".join(chr(c) for c in TARGET if c != 32)


def build(font, bold):
    b = "<w:b/>" if bold else ""
    r = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>{b}<w:sz w:val="40"/>'
    body = (f'<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr>'
            f'<w:t xml:space="preserve">{pg.esc(probe_text())}</w:t></w:r></w:p>')
    pgsz = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    mar = ('<w:pgMar w:top="720" w:right="720" w:bottom="720" '
           'w:left="720" w:header="708" w:footer="708" w:gutter="0"/>')
    return pg.doc(body + pg.sectpr(pgsz=pgsz, mar=mar, grid=''))


def run():
    import win32com.client, fitz
    from fontTools.ttLib import TTFont
    os.makedirs(OUTDIR, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    entries = []
    try:
        for label, spec in FACES.items():
            nm = label.replace(" ", "_")
            src = os.path.abspath(os.path.join(OUTDIR, nm + ".docx"))
            pdf = os.path.abspath(os.path.join(OUTDIR, nm + ".pdf"))
            pg.write_docx(src, build(spec["font"], spec["bold"]),
                          font=spec["font"], sz="40", compat="15", cpunct=False)
            doc = word.Documents.Open(src, ReadOnly=True)
            try:
                doc.ExportAsFixedFormat(pdf, 17)
            finally:
                doc.Close(False)
            d = fitz.open(pdf)
            # pick the embedded font whose basename matches the face
            want = spec["font"].replace(" ", "")
            got = None
            for xref, ext, ftype, name, refname, enc in d.get_page_fonts(0):
                base = name.split("+")[-1]
                flat = base.replace(" ", "")
                is_bold = ",Bold" in base or "Bold" in flat.replace(want, "")
                if flat.startswith(want) and (is_bold == spec["bold"]):
                    got = (xref, name)
                    break
            if not got:
                print(f"  {label}: NO embedded font matched "
                      f"({[n for _,_,_,n,_,_ in d.get_page_fonts(0)]})")
                d.close()
                continue
            buf = d.extract_font(got[0])[3]
            ttf = os.path.join(OUTDIR, nm + ".ttf")
            open(ttf, "wb").write(buf)
            d.close()

            f = TTFont(ttf)
            upm = f["head"].unitsPerEm
            hhea, os2 = f["hhea"], f["OS/2"]
            cmap, hmtx = f.getBestCmap(), f["hmtx"]
            widths = {}
            for c in TARGET:
                g = cmap.get(c)
                if g and g in hmtx.metrics:
                    widths[str(c)] = hmtx[g][0]
            e = dict(family=label, units_per_em=upm,
                     ascender=hhea.ascent, descender=hhea.descent,
                     line_gap=hhea.lineGap,
                     win_ascent=os2.usWinAscent, win_descent=os2.usWinDescent,
                     typo_ascender=os2.sTypoAscender,
                     typo_descender=os2.sTypoDescender,
                     typo_line_gap=os2.sTypoLineGap,
                     widths=widths)
            entries.append(e)
            miss = [c for c in TARGET if str(c) not in widths]
            print(f"  {label}: embedded={got[1]} upm={upm} "
                  f"hhea={hhea.ascent}/{hhea.descent}/{hhea.lineGap} "
                  f"win={os2.usWinAscent}/{os2.usWinDescent} "
                  f"covered={len(widths)}/{len(TARGET)} missing={miss[:8]}")
    finally:
        word.Quit()
    out = os.path.join(OUTDIR, "_entries.json")
    json.dump(entries, open(out, "w", encoding="utf-8"), indent=1)
    print("\nwrote", out)
    return entries


if __name__ == "__main__":
    ents = run()
    if "--emit" in sys.argv:
        print(json.dumps(ents, indent=1))
