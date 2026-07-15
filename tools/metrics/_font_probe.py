"""What does Word ACTUALLY render for a given font name? (generalises S866)

For any font family name, author a probe docx that sets rFonts ascii/hAnsi to
it over every target codepoint, export via Word COM to PDF, and read back the
PDF span's REAL font name. Word tells us one of two things:

  * the SAME family  -> Word has the font (installed, or Office-internal like
    Aptos which exists on no disk path). The PDF embeds a subset -> extract it
    and read real metrics with fontTools (the S866 route).
  * a DIFFERENT family -> Word SUBSTITUTED it (PANOSE/font-linking). Then the
    fix is an ALIAS to the substitute (the S796 Humnst777->Calibri pattern),
    NOT a metrics table for a font nobody has.

This removes the guesswork from the "unknown font" lever (S819 Segoe UI, S846
Helvetica, S855 Arial Narrow, S858 SimSun, S860 Bookman, S866 Aptos): every
name gets a measured answer, including for uninstalled/commercial faces.

Usage:
  python _font_probe.py "Amasis MT Pro Black" "Minion Pro" ...
  python _font_probe.py --fails      # the unhandled names in EN Phase-1 FAILs
Artifacts + extracted ttf/metrics: pipeline_data/_font_probe/
"""
import os, sys, json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _probe_gen as pg

OUTDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                      "pipeline_data", "_font_probe")

TARGET = ([c for c in range(32, 127)]
          + [8216, 8217, 8220, 8221]
          + list(range(9312, 9332)))

DEFAULT_FAILS = ["Amasis MT Pro Black", "Amasis MT Std Black", "Calibri Light",
                 "Cambria Math", "Minion Pro", "Times", "Trebuchet MS",
                 "Aharoni", "Grammarsaurus"]


def probe_text():
    return " ".join(chr(c) for c in TARGET if c != 32)


def build(font):
    r = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="40"/>'
    body = (f'<w:p><w:pPr><w:jc w:val="left"/></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr>'
            f'<w:t xml:space="preserve">{pg.esc(probe_text())}</w:t></w:r></w:p>')
    pgsz = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    mar = ('<w:pgMar w:top="720" w:right="720" w:bottom="720" '
           'w:left="720" w:header="708" w:footer="708" w:gutter="0"/>')
    return pg.doc(body + pg.sectpr(pgsz=pgsz, mar=mar, grid=''))


def slug(name):
    return "".join(ch if ch.isalnum() else "_" for ch in name)


def run(names):
    import win32com.client, fitz
    from fontTools.ttLib import TTFont
    os.makedirs(OUTDIR, exist_ok=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    report = {}
    try:
        for name in names:
            nm = slug(name)
            src = os.path.abspath(os.path.join(OUTDIR, nm + ".docx"))
            pdf = os.path.abspath(os.path.join(OUTDIR, nm + ".pdf"))
            pg.write_docx(src, build(name), font=name, sz="40",
                          compat="15", cpunct=False)
            doc = word.Documents.Open(src, ReadOnly=True)
            try:
                doc.ExportAsFixedFormat(pdf, 17)
            finally:
                doc.Close(False)
            d = fitz.open(pdf)
            # dominant span font on page 1
            from collections import Counter
            c = Counter()
            for blk in d[0].get_text("dict")["blocks"]:
                for ln in blk.get("lines", []):
                    for s in ln["spans"]:
                        c[s["font"]] += len(s["text"])
            rendered = c.most_common(1)[0][0] if c else None
            base_req = name.replace(" ", "").lower()
            base_got = (rendered or "").split("+")[-1].split(",")[0].replace(" ", "").lower()
            same = base_got.startswith(base_req) or base_req.startswith(base_got)
            entry = dict(requested=name, rendered=rendered,
                         verdict="HAS FONT" if same else "SUBSTITUTED",
                         spans=dict(c.most_common(4)))
            # if Word has it, extract the embedded subset -> real metrics
            if same:
                for xref, ext, ftype, fname, refname, enc in d.get_page_fonts(0):
                    if fname.split("+")[-1].split(",")[0].replace(" ", "").lower() \
                            .startswith(base_req):
                        buf = d.extract_font(xref)[3]
                        ttf = os.path.join(OUTDIR, nm + ".ttf")
                        open(ttf, "wb").write(buf)
                        f = TTFont(ttf)
                        upm = f["head"].unitsPerEm
                        hh, o2 = f["hhea"], f["OS/2"]
                        cm, hm = f.getBestCmap(), f["hmtx"]
                        widths = {str(cp): hm[cm[cp]][0]
                                  for cp in TARGET if cm.get(cp) in hm.metrics}
                        entry["metrics"] = dict(
                            family=name, units_per_em=upm,
                            ascender=hh.ascent, descender=hh.descent,
                            line_gap=hh.lineGap,
                            win_ascent=o2.usWinAscent, win_descent=o2.usWinDescent,
                            typo_ascender=o2.sTypoAscender,
                            typo_descender=o2.sTypoDescender,
                            typo_line_gap=o2.sTypoLineGap, widths=widths)
                        entry["hhea_line_em"] = round(
                            (hh.ascent - hh.descent + hh.lineGap) / upm, 4)
                        entry["covered"] = len(widths)
                        break
            d.close()
            report[name] = entry
            v = entry["verdict"]
            extra = (f" line={entry.get('hhea_line_em')}em "
                     f"covered={entry.get('covered')}" if "metrics" in entry else "")
            print(f"  {name!r:28} -> {v:12} rendered={rendered!r}{extra}", flush=True)
    finally:
        word.Quit()
    out = os.path.join(OUTDIR, "_report.json")
    json.dump(report, open(out, "w", encoding="utf-8"), indent=1, ensure_ascii=False)
    print("\nwrote", out)
    return report


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--fails"]
    run(args if args else DEFAULT_FAILS)
