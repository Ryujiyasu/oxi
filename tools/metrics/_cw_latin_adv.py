# -*- coding: utf-8 -*-
"""What advance does Word give a LATIN glyph inside a JAPANESE document?

S672 emits pure-Latin LINES at the true em and LATINEM breaks pure-Latin
DOCUMENTS at it, so both are scoped away from the case that is still wrong: a
Latin run sitting in a CJK line of a CJK document. There Oxi keeps the 10tw
(0.5pt) rounded width for the break AND the render, and c7b923e5 p2 shows the
cost -- Word draws "API" 18.00 wide, Oxi 18.76, and 45 characters of that
accumulate into a two-character line-length error.

This is `_cb_pgothic_adv.py`'s in-span pair method turned toward Latin, with
one correction that matters more here than it did there: the advance is read
from the char's **origin** (the pen), not its bbox. A bbox is ink, and Latin
side bearings differ per glyph, so bbox deltas mix the advance with the
letterform. CJK fullwidth glyphs hid that; 'I' next to 'A' does not.

Every Latin-Latin pair inside one span of a ragged (paragraph-final, hence
unjustified) line is collected, and Word's number is put beside the three Oxi
could return for it: the design em, that em rounded to 10tw (what
`char_width_pt_with_gdi_map` actually returns), and the com_tw override.

    python _cw_latin_adv.py                  # c7b923e5
    python _cw_latin_adv.py 04b88e 34140b    # several documents
"""
import collections
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import _cb_budget as B  # noqa: E402

DATA = os.path.join(B.REPO, "crates", "oxidocs-core", "src", "font", "data")
METRICS = os.path.join(DATA, "font_metrics_compact.json")
COM_TW = os.path.join(DATA, "com_tw_overrides.json")
RAGGED_MARGIN = 12.0     # a line ending this short of the widest is unjustified


def is_latin(ch):
    cp = ord(ch)
    return (0x21 <= cp <= 0x7E) or (0xA1 <= cp <= 0x24F)


def is_cjk(ch):
    cp = ord(ch)
    return (0x3000 <= cp <= 0x30FF or 0x4E00 <= cp <= 0x9FFF
            or 0xFF00 <= cp <= 0xFFEF or 0x3400 <= cp <= 0x4DBF)


# PDF font name -> the family key in the metrics table
def family_of(pdf_font):
    base = pdf_font.split("+")[-1]
    known = {
        "TimesNewRomanPSMT": "Times New Roman",
        "TimesNewRomanPS-BoldMT": "Times New Roman",
        "TimesNewRomanPS-ItalicMT": "Times New Roman",
        "ArialMT": "Arial", "Arial-BoldMT": "Arial",
        "MS-PMincho": "MS PMincho", "MS-Mincho": "MS Mincho",
        "MS-PGothic": "MS PGothic", "MS-Gothic": "MS Gothic",
        "Century": "Century", "CenturyOldst": "Century",
        "Calibri": "Calibri", "Cambria": "Cambria",
        "MSPMincho": "MS PMincho", "MSMincho": "MS Mincho",
        "MSPGothic": "MS PGothic", "MSGothic": "MS Gothic",
    }
    if base in known:
        return known[base]
    return base.replace("-", " ")


def load_tables():
    metrics = {}
    for f in json.load(open(METRICS, encoding="utf-8")):
        metrics[f["family"]] = (f["units_per_em"], f.get("widths", {}))
    com = {}
    if os.path.exists(COM_TW):
        com = json.load(open(COM_TW, encoding="utf-8"))
    return metrics, com


def size_key(fs):
    return ("%.1f" % fs).rstrip("0").rstrip(".")


# ★Word's PDF span size is 600dpi-snapped and reads 0.57% high (sz=21 -> 10.56,
# sz=16 -> 8.04): n = round(pt*600/72), size = n*72/600. Ratios survive it,
# absolute pt do not -- and the 10tw grid is absolute, so the rounding has to be
# done at the NOMINAL size or it lands on the wrong step.
def nominal_size(pdf_size):
    for half_pt in range(2, 145):
        nom = half_pt / 2.0
        if abs(round(nom * 600.0 / 72.0) * 72.0 / 600.0 - pdf_size) < 0.005:
            return nom
    return pdf_size


def is_fullwidth(cp):
    return (0x1100 <= cp <= 0x115F or 0x2E80 <= cp <= 0xA4CF
            or 0xAC00 <= cp <= 0xD7A3 or 0xF900 <= cp <= 0xFAFF
            or 0xFE30 <= cp <= 0xFE6F or 0xFF00 <= cp <= 0xFF60
            or 0xFFE0 <= cp <= 0xFFE6)


def oxi_width(ch, fs, fam, upm, widths, ctab):
    """Mirror char_width_pt_with_gdi_map's branch order, so the column says
    what Oxi returns rather than what one of its branches would."""
    cp = ord(ch)
    pgothic = fam in ("MS PGothic", "MS PMincho", "HGPGothicM")
    if upm == 256 and not pgothic:
        if is_fullwidth(cp):
            return fs, "fullwidth"
        em = widths.get(str(cp))
        em = None if em is None else em / upm
        if em is not None and em <= 0.51:
            return fs / 2.0, "s546 half"
    ctw = ctab.get(str(cp))
    if ctw is not None:
        return ctw / 20.0, "com_tw"
    em = widths.get(str(cp))
    if em is not None:
        em /= upm
        return int(em * fs * 20.0 / 10.0 + 0.5) * 10.0 / 20.0, "10tw"
    return None, "-"


def collect(pdf_path):
    """(font, size, ch) -> [advance in pt] from unjustified lines only."""
    import fitz
    adv = collections.defaultdict(list)
    pdf = fitz.open(pdf_path)
    for pg in pdf:
        lines = []
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                txt = "".join(c["c"] for s in ln["spans"] for c in s["chars"])
                if txt.strip():
                    lines.append(ln)
        if not lines:
            continue
        edge = max(ln["bbox"][2] for ln in lines)
        for ln in lines:
            if ln["bbox"][2] > edge - RAGGED_MARGIN:
                continue                     # justified: advances are stretched
            for s in ln["spans"]:
                sc = s["chars"]
                for i in range(len(sc) - 1):
                    a, b2 = sc[i], sc[i + 1]
                    if not (is_latin(a["c"]) and (is_latin(b2["c"])
                                                  or b2["c"] == " ")):
                        continue
                    # ★the PEN, not the ink: Latin side bearings are per-glyph
                    d = b2["origin"][0] - a["origin"][0]
                    if not 0.05 * s["size"] < d < 1.6 * s["size"]:
                        continue
                    adv[(s["font"], round(s["size"], 2), a["c"])].append(d)
    return adv


def report(prefix, metrics, com):
    docx = B.docx_for(prefix)
    rt = docx[:-5] + "_rt.pdf"
    if not os.path.exists(rt):
        print("== %s ==  no Word PDF" % prefix)
        return []
    adv = collect(rt)
    # keep the busiest Latin (font,size) cohort of the document
    tally = collections.Counter()
    for (fnt, fs, ch), v in adv.items():
        tally[(fnt, fs)] += len(v)
    if not tally:
        print("== %s ==  no unjustified Latin pairs" % prefix)
        return []
    rows_all = []
    for (fnt, fs), n in tally.most_common(3):
        fam = family_of(fnt)
        upm, widths = metrics.get(fam, (None, {}))
        nom = nominal_size(fs)
        ctab = com.get(fam, {}).get(size_key(nom), {})
        print()
        print("== %s ==  %s  pdf %.2fpt -> nominal %.1fpt  (fam=%s upm=%s, n=%d)"
              % (os.path.basename(docx)[:30], fnt, fs, nom, fam, upm, n))
        print("%-3s %-4s %-8s %-8s %-8s %-9s %-8s %s"
              % ("ch", "n", "word_em", "des_em", "word_pt", "oxi_pt", "oxi-w", "via"))
        rows = []
        for (f2, s2, ch), v in adv.items():
            if (f2, s2) != (fnt, fs) or len(v) < 3:
                continue
            v = sorted(v)
            wem = v[len(v) // 2] / fs             # scale-free: the 600dpi snap cancels
            wpt = wem * nom                       # Word's advance at the nominal size
            de = widths.get(str(ord(ch)))
            dem = None if de is None or not upm else de / upm
            opt, via = (oxi_width(ch, nom, fam, upm, widths, ctab)
                        if upm else (None, "-"))
            rows.append((ch, len(v), wem, dem, wpt, opt, via))
        rows.sort(key=lambda r: -r[1])
        for ch, n2, wem, dem, wpt, opt, via in rows[:24]:
            print("%-3s %-4d %-8.4f %-8s %-8.3f %-9s %-8s %s"
                  % (ch, n2, wem,
                     "%.4f" % dem if dem else "-", wpt,
                     "%.3f" % opt if opt else "-",
                     "%+.3f" % (opt - wpt) if opt else "-", via))
        if rows:
            dd = [dem - wem for _c, _n, wem, dem, _w, _o, _v in rows if dem]
            oo = [o - w for _c, _n, _e, _d, w, o, _v in rows if o]
            print("  design_em-vs-Word  mean %+.4f em (max |%.4f|)  n=%d"
                  % (sum(dd) / len(dd), max(abs(x) for x in dd), len(dd)))
            if oo:
                print("  ★OXI-vs-Word      mean %+.3f pt (max |%.3f|)  n=%d"
                      % (sum(oo) / len(oo), max(abs(x) for x in oo), len(oo)))
        rows_all.append((prefix, fnt, nom, rows))
    return rows_all


def main():
    prefixes = sys.argv[1:] or ["c7b923e5"]
    metrics, com = load_tables()
    for p in prefixes:
        report(p, metrics, com)


if __name__ == "__main__":
    main()
